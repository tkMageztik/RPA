using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Conditions;
using FlaUI.Core.Definitions;
using FlaUI.Core.Input;
using FlaUI.Core.Patterns;
using FlaUI.Core.Tools;
using FlaUI.UIA3;

namespace NS.RPA.Demo
{
    class Program
    {
        // Generic UWP host process — UWP apps are rendered inside this Win32 host.
        private const string UwpHostProcessName = "ApplicationFrameHost";

        // Pre-built indent strings to avoid per-call allocations during deep tree traversal.
        private static readonly string[] IndentCache = BuildIndentCache(20);
        private static string[] BuildIndentCache(int size)
        {
            var cache = new string[size];
            for (int i = 0; i < size; i++) cache[i] = new string(' ', i * 2);
            return cache;
        }
        static string Indent(int level) => level < IndentCache.Length ? IndentCache[level] : new string(' ', level * 2);

        // ─── Logging ─────────────────────────────────────────────────────────────────

        static void Log(string message)
        {
            Console.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] [INFO ] {message}");
        }

        static void LogError(string message, Exception? ex = null)
        {
            string ts = DateTime.Now.ToString("HH:mm:ss.fff");
            Console.Error.WriteLine($"[{ts}] [ERROR] {message}");
            if (ex != null)
            {
                Console.Error.WriteLine($"[{ts}] [ERROR] Exception type : {ex.GetType().FullName}");
                Console.Error.WriteLine($"[{ts}] [ERROR] Message        : {ex.Message}");
                Console.Error.WriteLine($"[{ts}] [ERROR] Stack trace    :\n{ex.StackTrace}");
                if (ex.InnerException != null)
                    Console.Error.WriteLine($"[{ts}] [ERROR] Inner exception: {ex.InnerException.Message}");
            }
        }

        // ─── Entry point ─────────────────────────────────────────────────────────────

        static int Main(string[] args)
        {
            try
            {
                string appPath = "calc.exe";
                bool pocMode = false;
                bool contextMode = false;
                bool clickFirst = false;
                string? outputFile = null;
                string? windowTitle = null;
                string? clickByName = null;
                string? clickById = null;
                string? clickAtCoords = null;

                for (int i = 0; i < args.Length; i++)
                {
                    if (args[i].Equals("--poc", StringComparison.OrdinalIgnoreCase))
                        pocMode = true;
                    else if (args[i].Equals("--context", StringComparison.OrdinalIgnoreCase))
                        contextMode = true;
                    else if (args[i].Equals("--click-first", StringComparison.OrdinalIgnoreCase))
                        clickFirst = true;
                    else if (args[i].Equals("--out", StringComparison.OrdinalIgnoreCase) && i + 1 < args.Length)
                        outputFile = args[++i];
                    else if (args[i].Equals("--title", StringComparison.OrdinalIgnoreCase) && i + 1 < args.Length)
                        windowTitle = args[++i];
                    else if (args[i].Equals("--click-name", StringComparison.OrdinalIgnoreCase) && i + 1 < args.Length)
                        clickByName = args[++i];
                    else if (args[i].Equals("--click-id", StringComparison.OrdinalIgnoreCase) && i + 1 < args.Length)
                        clickById = args[++i];
                    else if (args[i].Equals("--click-at", StringComparison.OrdinalIgnoreCase) && i + 1 < args.Length)
                        clickAtCoords = args[++i];
                    else
                        appPath = args[i];
                }

                if (string.IsNullOrWhiteSpace(appPath))
                    throw new ArgumentException("Application path must not be empty.");

                string mode = contextMode ? "Context Menu Explorer" : pocMode ? "POC Inspector" : "Demo (keyboard)";
                Log($"Mode        : {mode}");
                Log($"Application : {appPath}");
                if (windowTitle  != null) Log($"Window title: '{windowTitle}' (override)");
                if (outputFile   != null) Log($"Output file : {outputFile}");
                if (clickByName  != null) Log($"Click name  : '{clickByName}'");
                if (clickById    != null) Log($"Click id    : '{clickById}'");
                if (clickAtCoords != null) Log($"Click at    : {clickAtCoords}");

                using var automation = new UIA3Automation();

                var (_, mainWindow) = GetOrLaunchApp(appPath, automation, windowTitle);

                if (contextMode)
                    RunContextMenuExplorer(mainWindow, automation, clickByName, clickById, clickAtCoords, clickFirst, outputFile);
                else if (pocMode)
                    RunPocInspector(mainWindow, outputFile);
                else
                    RunCalcDemo(mainWindow);

                return 0;
            }
            catch (Exception ex)
            {
                LogError("Unhandled exception — automation aborted.", ex);
                return 1;
            }
        }

        // ─── App connection ───────────────────────────────────────────────────────────

        /// <summary>
        /// Attaches to an already-running instance of the target application, or launches it from
        /// the given path and waits reactively until its main window is ready.
        /// </summary>
        /// <param name="appPath">Full path or bare executable name.</param>
        /// <param name="automation">The UIA3 automation instance.</param>
        /// <param name="windowTitle">Optional window title override (substring match) for locating the window.</param>
        /// <returns>A tuple containing the attached <see cref="Application"/> and its ready <see cref="Window"/> main window.</returns>
        static (Application app, Window mainWindow) GetOrLaunchApp(
            string appPath, UIA3Automation automation, string? windowTitle = null)
        {
            string processName = Path.GetFileNameWithoutExtension(appPath);
            string searchLabel = windowTitle ?? processName;
            Log($"Searching for existing window of '{searchLabel}' using multi-strategy approach...");

            var existingWindow = FindWindowMultiStrategy(processName, automation, windowTitle);
            if (existingWindow != null)
            {
                Log("Application is already running — attaching to existing window.");
                var existingProcess = Process.GetProcessById((int)existingWindow.Properties.ProcessId.Value);
                return (Application.Attach(existingProcess), existingWindow);
            }

            Log($"No existing window found — launching '{appPath}'...");
            try
            {
                Process.Start(appPath);
                Log($"Launch command issued for '{appPath}'.");
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException(
                    $"Failed to launch '{appPath}'. Verify the path is correct and accessible.", ex);
            }

            Log($"Waiting for window of '{searchLabel}' to appear (multi-strategy)...");
            var retryResult = Retry.WhileNull(
                () => FindWindowMultiStrategy(processName, automation, windowTitle),
                TimeSpan.FromSeconds(30),
                throwOnTimeout: true,
                timeoutMessage: $"Window for '{searchLabel}' did not appear within 30 seconds.");

            var mainWindow = retryResult.Result!;
            var launchedProcess = Process.GetProcessById((int)mainWindow.Properties.ProcessId.Value);
            Log($"Window is ready — PID {launchedProcess.Id}, ProcessName '{launchedProcess.ProcessName}'.");

            return (Application.Attach(launchedProcess), mainWindow);
        }

        /// <summary>
        /// Finds the main window of a target application using three strategies in order.
        /// Returns null if no matching window is found on this attempt (caller should retry).
        /// </summary>
        static Window? FindWindowMultiStrategy(
            string processName, UIA3Automation automation, string? windowTitle = null)
        {
            // ── Strategy 1: Exact process name ─────────────────────────────────────────
            var procs = Process.GetProcessesByName(processName);
            foreach (var proc in procs)
            {
                int pid = -1;
                try
                {
                    pid = proc.Id;
                    var app = Application.Attach(proc);
                    var win = app.GetMainWindow(automation);
                    if (win != null && MatchesTitle(win, windowTitle))
                    {
                        Log($"[Strategy 1 - Exact process] Found '{processName}' (PID {pid}).");
                        DetectAndLogAppType(win);
                        return win;
                    }
                }
                catch (Exception ex) { Log($"[Strategy 1] Skipping PID {pid}: {ex.Message}"); }
            }

            // ── Strategy 2: UWP ApplicationFrameHost ──────────────────────────────────
            var hostProcs = Process.GetProcessesByName(UwpHostProcessName);
            foreach (var hostProc in hostProcs)
            {
                int hostPid = -1;
                try
                {
                    hostPid = hostProc.Id;
                    var hostApp = Application.Attach(hostProc);
                    foreach (var win in hostApp.GetAllTopLevelWindows(automation))
                    {
                        string title = win.Title ?? "";
                        string className = win.Properties.ClassName.IsSupported
                            ? (win.Properties.ClassName.Value ?? "") : "";

                        bool titleHint = windowTitle != null
                            ? title.IndexOf(windowTitle, StringComparison.OrdinalIgnoreCase) >= 0
                            : title.IndexOf(processName, StringComparison.OrdinalIgnoreCase) >= 0
                              || className.IndexOf(processName, StringComparison.OrdinalIgnoreCase) >= 0;

                        if (titleHint)
                        {
                            Log($"[Strategy 2 - UWP host] Found in '{UwpHostProcessName}' — Title: '{title}'.");
                            DetectAndLogAppType(win);
                            return win;
                        }
                    }
                }
                catch (Exception ex) { Log($"[Strategy 2] Skipping UWP host PID {hostPid}: {ex.Message}"); }
            }

            // ── Strategy 3: Full desktop scan ─────────────────────────────────────────
            try
            {
                var desktop = automation.GetDesktop();
                foreach (var child in desktop.FindAllChildren())
                {
                    try
                    {
                        var win = child.AsWindow();
                        if (win == null) continue;

                        var pid = (int)win.Properties.ProcessId.Value;
                        var proc = Process.GetProcessById(pid);

                        bool match = proc.ProcessName.Equals(processName, StringComparison.OrdinalIgnoreCase)
                                     && MatchesTitle(win, windowTitle);
                        if (match)
                        {
                            Log($"[Strategy 3 - Desktop scan] Found window owned by '{proc.ProcessName}' (PID {pid}).");
                            DetectAndLogAppType(win);
                            return win;
                        }
                    }
                    catch (Exception ex) { Log($"[Strategy 3] Skipping desktop child: {ex.Message}"); }
                }
            }
            catch (Exception ex) { Log($"[Strategy 3] Desktop scan failed: {ex.Message}"); }

            return null;
        }

        static bool MatchesTitle(Window win, string? windowTitle)
        {
            if (windowTitle == null) return true;
            string title = win.Title ?? "";
            return title.IndexOf(windowTitle, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        /// <summary>
        /// Reads the window ClassName to detect and log the type of application framework.
        /// </summary>
        static string DetectAppType(Window window)
        {
            try
            {
                string className = window.Properties.ClassName.IsSupported
                    ? (window.Properties.ClassName.Value ?? "") : "";

                if (className.Contains("ApplicationFrameWindow") || className.Contains("Windows.UI.Core"))
                    return "UWP (Modern Windows App)";
                if (className.StartsWith("pbframe", StringComparison.OrdinalIgnoreCase)
                    || className.StartsWith("PBGUI", StringComparison.OrdinalIgnoreCase)
                    || className.StartsWith("PBGUIObject", StringComparison.OrdinalIgnoreCase))
                    return "PowerBuilder";
                if (className.Contains("WindowsForms"))
                    return "Windows Forms (.NET)";
                if (className.Contains("HwndWrapper") || className.Contains("HwndSource"))
                    return "WPF (.NET)";
                if (className.StartsWith("Chrome_", StringComparison.OrdinalIgnoreCase)
                    || className.Contains("Electron"))
                    return "Electron / Chrome-based";
                return "Win32 / Classic";
            }
            catch { return "Unknown"; }
        }

        static void DetectAndLogAppType(Window window)
        {
            string appType = DetectAppType(window);
            string className = "";
            try { className = window.Properties.ClassName.Value ?? ""; }
            catch (Exception ex) { Log($"Could not read ClassName: {ex.Message}"); }
            Log($"Detected app type: {appType}");
            Log($"Window ClassName : '{className}'");
        }

        // ─── Demo mode (keyboard input) ───────────────────────────────────────────────

        /// <summary>
        /// Focuses the main window and types 1234 + 345, then presses Enter.
        /// </summary>
        /// <param name="mainWindow">The ready main window element returned by <see cref="GetOrLaunchApp"/>.</param>
        static void RunCalcDemo(Window mainWindow)
        {
            Log("--- Starting keyboard demo ---");
            Log($"Focusing window '{mainWindow.Title}'...");
            mainWindow.Focus();
            Log("Window focused.");

            Log("Typing '1234'...");
            Keyboard.Type("1234");
            Log("Typed '1234'.");

            Log("Typing '+'...");
            Keyboard.Type("+");
            Log("Typed '+'.");

            Log("Typing '345'...");
            Keyboard.Type("345");
            Log("Typed '345'.");

            Log("Pressing Enter...");
            Keyboard.Press(FlaUI.Core.WindowsAPI.VirtualKeyShort.RETURN);
            Log("Enter pressed.");

            Log("--- Keyboard demo completed: 1234 + 345 ---");
            Console.WriteLine();
            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }

        // ─── POC Inspector mode ───────────────────────────────────────────────────────

        /// <summary>
        /// Walks the UI Automation element tree and dumps full element details to console
        /// (and optionally to a file). The output contains everything needed to write
        /// FlaUI automation code for any control in the target application.
        /// </summary>
        /// <param name="mainWindow">The ready main window element returned by <see cref="GetOrLaunchApp"/>.</param>
        /// <param name="outputFile">Optional file path to save the report. If null, console only.</param>
        static void RunPocInspector(Window mainWindow, string? outputFile = null)
        {
            string appType = DetectAppType(mainWindow);
            string windowClass = "";
            try { windowClass = mainWindow.Properties.ClassName.Value ?? ""; }
            catch (Exception ex) { Log($"Could not read ClassName: {ex.Message}"); }

            // If a file path was given, install a TeeWriter so every byte written to
            // Console.Out is simultaneously written to the file — console and file are identical.
            TextWriter originalOut = Console.Out;
            StreamWriter? fileWriter = null;
            if (outputFile != null)
            {
                try
                {
                    fileWriter = new StreamWriter(outputFile, append: false, Encoding.UTF8)
                    {
                        AutoFlush = true
                    };
                    Console.SetOut(new TeeWriter(originalOut, fileWriter));
                    Log($"Tee output active — mirroring console to: {outputFile}");
                }
                catch (Exception ex)
                {
                    LogError($"Could not open output file '{outputFile}' — console only.", ex);
                }
            }

            try
            {
                Log("--- Starting POC UI Inspector ---");
                Log("Output contains: ControlType | Name | AutomationId | ClassName | BoundingRect |");
                Log("                 IsEnabled | IsOffscreen | SupportedPatterns | CurrentValue");
                Log("FlaUI code hints are shown for each interactable element.");

                var sb = new StringBuilder();
                sb.AppendLine("╔══════════════════════════════════════════════════════════════════════════════╗");
                sb.AppendLine("║              NS.RPA.Demo — POC UI Automation Inspector Report               ║");
                sb.AppendLine("╚══════════════════════════════════════════════════════════════════════════════╝");
                sb.AppendLine();
                sb.AppendLine($"Generated    : {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                sb.AppendLine($"Window Title : {mainWindow.Title}");
                sb.AppendLine($"Window Class : {windowClass}");
                sb.AppendLine($"App Type     : {appType}");
                sb.AppendLine();
                sb.AppendLine("── Legend ───────────────────────────────────────────────────────────────────");
                sb.AppendLine("  [BUTTON]   → element.AsButton().Invoke()  or  Keyboard focus + Enter");
                sb.AppendLine("  [EDIT]     → element.AsTextBox().Enter(\"value\")  or  ValuePattern.SetValue");
                sb.AppendLine("  [CHECK]    → element.AsCheckBox().Toggle()");
                sb.AppendLine("  [COMBO]    → element.AsComboBox().Select(\"item\")");
                sb.AppendLine("  [LIST]     → element.AsListBox().Select(\"item\")");
                sb.AppendLine("  [MENU]     → element.AsMenuItem().Invoke()  or  ExpandCollapse");
                sb.AppendLine("  [TREE]     → ExpandCollapsePattern.Expand() / Collapse()");
                sb.AppendLine("  [COORD]    → Mouse.Click(new Point(cx,cy))  — use when element has no pattern");
                sb.AppendLine("  [OPAQUE]   → Not accessible via UIA (e.g. DataWindow). Use coords or WinAPI.");
                sb.AppendLine();
                sb.AppendLine("── UI Element Tree ──────────────────────────────────────────────────────────");

                int nodeCount = PrintRichUiaTree(mainWindow, sb, indent: 0, maxDepth: 10);

                sb.AppendLine();
                sb.AppendLine($"── Summary: {nodeCount} element(s) found ─────────────────────────────────────");
                if (appType == "PowerBuilder")
                {
                    sb.AppendLine();
                    sb.AppendLine("PowerBuilder note:");
                    sb.AppendLine("  Controls showing [OPAQUE] are DataWindows or custom PB painters.");
                    sb.AppendLine("  For those, use coordinate-based Mouse.Click(new Point(x,y)) with the");
                    sb.AppendLine("  BoundingRect center coordinates shown above each [OPAQUE] element.");
                }

                // Console.WriteLine goes through TeeWriter, so the file gets exactly this too.
                Console.WriteLine();
                Console.WriteLine(sb.ToString());

                Log($"--- POC inspection completed — {nodeCount} element(s) found ---");
                if (outputFile != null && fileWriter != null)
                    Log($"Report saved to: {outputFile}");  // logged while TeeWriter is still active → goes to file too
            }
            finally
            {
                // Always restore Console.Out and close the file, even if an exception occurred.
                Console.SetOut(originalOut);
                fileWriter?.Dispose();
            }

            Console.WriteLine();
            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }

        /// <summary>
        /// Recursively walks the UIA element tree and appends rich per-element details to the
        /// StringBuilder, including interaction patterns and FlaUI code hints.
        /// </summary>
        static int PrintRichUiaTree(AutomationElement element, StringBuilder sb, int indent, int maxDepth)
        {
            if (indent > maxDepth)
            {
                sb.AppendLine(Indent(indent) + "... (max depth reached)");
                return 0;
            }

            try
            {
                // ── Core identity ──────────────────────────────────────────────────────
                string name         = SafeGet(() => element.Properties.Name.Value ?? "");
                string automationId = SafeGet(() => element.Properties.AutomationId.Value ?? "");
                string className    = SafeGet(() => element.Properties.ClassName.Value ?? "");
                string controlType  = SafeGet(() => element.Properties.ControlType.Value.ToString(), "?");
                bool   isEnabled    = SafeGet(() => element.Properties.IsEnabled.Value, false);
                bool   isOffscreen  = SafeGet(() => element.Properties.IsOffscreen.Value, true);

                // ── Bounding rectangle ─────────────────────────────────────────────────
                var rect = SafeGet(() => element.Properties.BoundingRectangle.Value,
                                   new System.Drawing.Rectangle());
                int cx = rect.X + rect.Width  / 2;
                int cy = rect.Y + rect.Height / 2;
                string rectStr = rect.IsEmpty
                    ? "n/a"
                    : $"X={rect.X} Y={rect.Y} W={rect.Width} H={rect.Height} (center {cx},{cy})";

                // ── Supported patterns and current values ──────────────────────────────
                var patterns = new List<string>();
                string currentValue = "";
                string flaUiHint    = "";

                if (element.Patterns.Invoke.IsSupported)
                {
                    patterns.Add("Invoke");
                    flaUiHint = $"element.Patterns.Invoke.Pattern.Invoke();  // click this {controlType}";
                }
                if (element.Patterns.Value.IsSupported)
                {
                    patterns.Add("Value");
                    currentValue = SafeGet(() => element.Patterns.Value.Pattern.Value.Value ?? "", "");
                    if (string.IsNullOrEmpty(flaUiHint))
                        flaUiHint = $"element.Patterns.Value.Pattern.SetValue(\"newText\");  // set text in this {controlType}";
                }
                if (element.Patterns.Toggle.IsSupported)
                {
                    patterns.Add("Toggle");
                    string toggleState = SafeGet(() => element.Patterns.Toggle.Pattern.ToggleState.Value.ToString(), "?");
                    currentValue = $"ToggleState={toggleState}";
                    if (string.IsNullOrEmpty(flaUiHint))
                        flaUiHint = $"element.Patterns.Toggle.Pattern.Toggle();  // check/uncheck";
                }
                if (element.Patterns.SelectionItem.IsSupported)
                {
                    patterns.Add("SelectionItem");
                    bool isSelected = SafeGet(() => element.Patterns.SelectionItem.Pattern.IsSelected.Value, false);
                    currentValue = $"IsSelected={isSelected}";
                    if (string.IsNullOrEmpty(flaUiHint))
                        flaUiHint = $"element.Patterns.SelectionItem.Pattern.Select();  // select this item";
                }
                if (element.Patterns.Selection.IsSupported)
                {
                    patterns.Add("Selection");
                    if (string.IsNullOrEmpty(flaUiHint))
                        flaUiHint = $"element.AsComboBox().Select(\"itemText\");  // or .AsListBox()";
                }
                if (element.Patterns.ExpandCollapse.IsSupported)
                {
                    patterns.Add("ExpandCollapse");
                    string ecState = SafeGet(() => element.Patterns.ExpandCollapse.Pattern.ExpandCollapseState.Value.ToString(), "?");
                    currentValue = $"ExpandCollapseState={ecState}";
                    if (string.IsNullOrEmpty(flaUiHint))
                        flaUiHint = $"element.Patterns.ExpandCollapse.Pattern.Expand();  // open menu/tree";
                }
                if (element.Patterns.RangeValue.IsSupported)
                {
                    patterns.Add("RangeValue");
                    string rv = SafeGet(() => element.Patterns.RangeValue.Pattern.Value.Value.ToString("F2"), "?");
                    currentValue = $"RangeValue={rv}";
                    if (string.IsNullOrEmpty(flaUiHint))
                        flaUiHint = $"element.Patterns.RangeValue.Pattern.SetValue(0.5);  // slider/progress";
                }
                if (element.Patterns.Scroll.IsSupported)
                    patterns.Add("Scroll");
                if (element.Patterns.Text.IsSupported)
                    patterns.Add("Text");
                if (element.Patterns.Window.IsSupported)
                    patterns.Add("Window");

                bool hasNoPatterns = patterns.Count == 0;
                bool isOpaque = hasNoPatterns && !isEnabled;
                string tag = isOpaque ? "[OPAQUE]" : $"[{controlType.ToUpperInvariant()}]";

                // ── Format the entry ───────────────────────────────────────────────────
                sb.AppendLine($"{Indent(indent)}{tag}");
                sb.AppendLine($"{Indent(indent)}  Name        : \"{name}\"");
                if (!string.IsNullOrEmpty(automationId))
                    sb.AppendLine($"{Indent(indent)}  AutomationId: \"{automationId}\"");
                if (!string.IsNullOrEmpty(className))
                    sb.AppendLine($"{Indent(indent)}  ClassName   : \"{className}\"");
                sb.AppendLine($"{Indent(indent)}  BoundingRect: {rectStr}");
                sb.AppendLine($"{Indent(indent)}  Enabled     : {isEnabled}  |  Offscreen: {isOffscreen}");
                if (patterns.Count > 0)
                    sb.AppendLine($"{Indent(indent)}  Patterns    : {string.Join(", ", patterns)}");
                if (!string.IsNullOrEmpty(currentValue))
                    sb.AppendLine($"{Indent(indent)}  Value       : {currentValue}");
                if (!string.IsNullOrEmpty(flaUiHint))
                    sb.AppendLine($"{Indent(indent)}  FlaUI hint  : {flaUiHint}");
                if (isOpaque)
                    sb.AppendLine($"{Indent(indent)}  >> OPAQUE: use Mouse.Click(new Point({cx},{cy})) or WinAPI");
                sb.AppendLine();

                // ── Recurse ────────────────────────────────────────────────────────────
                int count = 1;
                foreach (var child in element.FindAllChildren())
                    count += PrintRichUiaTree(child, sb, indent + 1, maxDepth);
                return count;
            }
            catch (Exception ex)
            {
                sb.AppendLine($"{Indent(indent)}[ERROR reading element at depth {indent}: {ex.Message}]");
                LogError($"Error reading UIA element at depth {indent}", ex);
                return 0;
            }
        }

        // ─── Context Menu Explorer mode ──────────────────────────────────────────────

        /// <summary>
        /// Clicks a target element (identified by name, AutomationId, or screen coordinates),
        /// then waits for and captures whatever popup or context menu appears as a result.
        /// The popup tree is dumped with full interaction details — ready to paste to an AI
        /// for automation help.  If <paramref name="clickFirst"/> is true, the first
        /// Invoke-capable item found in the popup is also invoked automatically.
        /// </summary>
        /// <param name="mainWindow">The main window of the target application.</param>
        /// <param name="automation">The UIA3 automation instance.</param>
        /// <param name="clickByName">Find and click the element whose Name contains this text (case-insensitive).</param>
        /// <param name="clickById">Find and click the element whose AutomationId equals this value.</param>
        /// <param name="clickAtCoords">Click at screen coordinates supplied as "x,y" (e.g. "1372,88").</param>
        /// <param name="clickFirst">If true, invoke the first Invoke-capable item found in the popup.</param>
        /// <param name="outputFile">Optional path to mirror output to a file (same TeeWriter mechanism as --poc).</param>
        static void RunContextMenuExplorer(
            Window mainWindow,
            UIA3Automation automation,
            string? clickByName,
            string? clickById,
            string? clickAtCoords,
            bool clickFirst,
            string? outputFile)
        {
            // ── Tee output to file if requested ──────────────────────────────────────
            TextWriter originalOut = Console.Out;
            StreamWriter? fileWriter = null;
            if (outputFile != null)
            {
                try
                {
                    fileWriter = new StreamWriter(outputFile, append: false, Encoding.UTF8) { AutoFlush = true };
                    Console.SetOut(new TeeWriter(originalOut, fileWriter));
                    Log($"Tee output active — mirroring console to: {outputFile}");
                }
                catch (Exception ex)
                {
                    LogError($"Could not open output file '{outputFile}' — console only.", ex);
                }
            }

            try
            {
                Log("--- Starting Context Menu Explorer ---");
                Log("Step 1: Snapshot existing top-level windows before clicking...");

                // Collect the runtime IDs of all current top-level windows so we can
                // detect any NEW window that appears after the click (Strategy A).
                var existingWindowIds = new HashSet<string>();
                try
                {
                    foreach (var child in automation.GetDesktop().FindAllChildren())
                    {
                        try
                        {
                            existingWindowIds.Add(child.Properties.RuntimeId.Value.ToString());
                        }
                        catch { /* skip inaccessible */ }
                    }
                }
                catch (Exception ex) { Log($"Snapshot warning: {ex.Message}"); }

                Log($"Snapshot complete — {existingWindowIds.Count} existing top-level window(s).");

                // Also snapshot the existing Invoke-capable descendants within the main window
                // so Strategy B can find NEW ones that appear after the click.
                var existingDescendantIds = new HashSet<string>();
                try
                {
                    foreach (var desc in mainWindow.FindAllDescendants())
                    {
                        try
                        {
                            if (desc.Patterns.Invoke.IsSupported)
                                existingDescendantIds.Add(desc.Properties.RuntimeId.Value.ToString());
                        }
                        catch { /* skip */ }
                    }
                }
                catch (Exception ex) { Log($"Descendant snapshot warning: {ex.Message}"); }

                Log($"Descendant snapshot complete — {existingDescendantIds.Count} existing Invoke-capable element(s).");

                // ── Step 2: Find and click the target element ─────────────────────────
                Log("Step 2: Finding and clicking the target element...");

                if (clickAtCoords != null)
                {
                    // Coordinate-based click — useful for unnamed/opaque elements like "..."
                    var parts = clickAtCoords.Split(',');
                    if (parts.Length != 2
                        || !int.TryParse(parts[0].Trim(), out int cx)
                        || !int.TryParse(parts[1].Trim(), out int cy))
                        throw new ArgumentException($"--click-at must be in format 'x,y', got: '{clickAtCoords}'");

                    Log($"Clicking at screen coordinates ({cx},{cy})...");
                    Mouse.Click(new System.Drawing.Point(cx, cy));
                    Log("Click sent.");
                }
                else
                {
                    // UIA element-based click
                    AutomationElement? target = null;

                    if (clickById != null)
                    {
                        Log($"Searching for element with AutomationId='{clickById}'...");
                        target = mainWindow.FindFirstDescendant(
                            cf => cf.ByAutomationId(clickById));
                    }
                    else if (clickByName != null)
                    {
                        Log($"Searching for element with Name containing '{clickByName}'...");
                        // FindAllDescendants and pick first case-insensitive match
                        foreach (var desc in mainWindow.FindAllDescendants())
                        {
                            try
                            {
                                string n = desc.Properties.Name.Value ?? "";
                                if (n.IndexOf(clickByName, StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    target = desc;
                                    break;
                                }
                            }
                            catch { /* skip */ }
                        }
                    }
                    else
                    {
                        throw new ArgumentException(
                            "Context mode requires --click-name, --click-id, or --click-at to identify the target element.");
                    }

                    if (target == null)
                        throw new InvalidOperationException(
                            $"No element found matching the given criteria. Run --poc first to inspect the tree.");

                    string targetName = SafeGet(() => target.Properties.Name.Value ?? "");
                    string targetId   = SafeGet(() => target.Properties.AutomationId.Value ?? "");
                    var    targetRect = SafeGet(() => target.Properties.BoundingRectangle.Value,
                                                new System.Drawing.Rectangle());
                    Log($"Target found — Name='{targetName}' AutomationId='{targetId}' " +
                        $"BoundingRect=X={targetRect.X} Y={targetRect.Y} W={targetRect.Width} H={targetRect.Height}");

                    Log("Clicking target element via Invoke pattern or mouse...");
                    if (target.Patterns.Invoke.IsSupported)
                    {
                        target.Patterns.Invoke.Pattern.Invoke();
                        Log("Invoked via InvokePattern.");
                    }
                    else
                    {
                        // Fall back to mouse click at center of bounding rect
                        int mx = targetRect.X + targetRect.Width  / 2;
                        int my = targetRect.Y + targetRect.Height / 2;
                        Mouse.Click(new System.Drawing.Point(mx, my));
                        Log($"Clicked at center ({mx},{my}) via mouse.");
                    }
                }

                // ── Step 3: Detect context menu / inline menu items ───────────────────
                // Windows-style context menus (right-click style) render as Menu/MenuItem
                // elements WITHIN the existing window tree — they are NOT separate top-level
                // windows. Give the menu a short time to render, then check inline first.
                Log("Step 3: Waiting 400 ms for context menu / inline items to render...");
                System.Threading.Thread.Sleep(400);

                AutomationElement? popup = null;

                // ── Strategy B (FIRST): inline Menu/MenuItem or new Invoke-capable items ──
                // Context menus that appear inline (like right-click on Windows desktop, or
                // WinUI CommandBarFlyout overflow) add Menu/MenuItem elements to the window
                // tree. Check here before wasting time looking for a new top-level window.
                Log("[Strategy B] Scanning for new Menu, MenuItem, or Invoke-capable elements...");
                AutomationElement? contextMenuRoot = null;
                var newItems = new List<AutomationElement>();
                try
                {
                    foreach (var desc in mainWindow.FindAllDescendants())
                    {
                        try
                        {
                            string rid = desc.Properties.RuntimeId.Value.ToString();
                            if (existingDescendantIds.Contains(rid)) continue;

                            // Match Menu/MenuItem control types (Windows context menus)
                            // and any other newly-appeared Invoke-capable visible element.
                            var ct = SafeGet(() => desc.Properties.ControlType.Value, ControlType.Unknown);
                            bool isMenuType = ct == ControlType.Menu || ct == ControlType.MenuItem;
                            bool isInvokable = desc.Patterns.Invoke.IsSupported
                                               && !desc.Properties.IsOffscreen.Value;

                            if (isMenuType || isInvokable)
                            {
                                // If we find a Menu container, use it as the root for the dump.
                                if (ct == ControlType.Menu && contextMenuRoot == null)
                                    contextMenuRoot = desc;
                                newItems.Add(desc);
                            }
                        }
                        catch { /* element gone or not ready — skip */ }
                    }
                }
                catch (Exception ex) { Log($"[Strategy B] Scan warning: {ex.Message}"); }

                if (newItems.Count > 0)
                {
                    Log($"[Strategy B] Found {newItems.Count} new element(s) in window tree — context menu is inline.");
                    // Use the Menu container if found, otherwise dump the whole main window.
                    popup = contextMenuRoot ?? mainWindow;
                }
                else
                {
                    // ── Strategy A (FALLBACK): new top-level window appeared ───────────────
                    // Some menus render in their own lightweight HWND (e.g. classic Win32
                    // context menus). Only check this if Strategy B found nothing.
                    Log("[Strategy B] No inline menu elements found.");
                    Log("[Strategy A] Checking for a new top-level window (up to 3 seconds)...");

                    var popupResult = Retry.WhileNull(
                        () =>
                        {
                            try
                            {
                                foreach (var child in automation.GetDesktop().FindAllChildren())
                                {
                                    try
                                    {
                                        string runtimeId = child.Properties.RuntimeId.Value.ToString();
                                        if (!existingWindowIds.Contains(runtimeId))
                                        {
                                            string title = SafeGet(() => child.AsWindow()?.Title ?? "");
                                            string cls   = SafeGet(() => child.Properties.ClassName.Value ?? "");
                                            Log($"[Strategy A] New top-level window — Title='{title}', Class='{cls}'");
                                            return child;
                                        }
                                    }
                                    catch { /* element gone */ }
                                }
                            }
                            catch { /* desktop scan failed */ }
                            return null;
                        },
                        TimeSpan.FromSeconds(3),
                        throwOnTimeout: false);

                    if (popupResult?.Result != null)
                    {
                        popup = popupResult.Result;
                        Log("[Strategy A] Context menu is in its own top-level window.");
                    }
                    else
                    {
                        Log("[Strategy A] No new window found either.");
                        Log("No context menu detected. Tips:");
                        Log("  1. Make sure the '...' button is actually being clicked (check coordinates).");
                        Log("  2. The menu may close before detection — try again.");
                        Log("  3. Some WinUI menus require ExpandCollapsePattern.Expand() instead of Invoke.");
                    }
                }

                // ── Step 4: Dump the popup tree ────────────────────────────────────────
                if (popup != null)
                {
                    Log("Step 4: Dumping popup / context menu tree...");
                    var sb = new StringBuilder();
                    sb.AppendLine();
                    sb.AppendLine("╔══════════════════════════════════════════════════════════════════════════════╗");
                    sb.AppendLine("║               Context Menu / Popup — Element Tree                          ║");
                    sb.AppendLine("╚══════════════════════════════════════════════════════════════════════════════╝");
                    sb.AppendLine($"Captured : {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                    sb.AppendLine();

                    int nodeCount = PrintRichUiaTree(popup, sb, indent: 0, maxDepth: 6);
                    Console.WriteLine(sb.ToString());
                    Log($"Popup tree: {nodeCount} element(s).");

                    // ── Step 5: Click first item if requested ──────────────────────────
                    if (clickFirst)
                    {
                        Log("Step 5: --click-first is set — invoking first Invoke-capable item in popup...");
                        AutomationElement? firstItem = null;
                        try
                        {
                            foreach (var desc in popup.FindAllDescendants())
                            {
                                try
                                {
                                    if (desc.Patterns.Invoke.IsSupported
                                        && !desc.Properties.IsOffscreen.Value
                                        && desc.Properties.IsEnabled.Value)
                                    {
                                        firstItem = desc;
                                        break;
                                    }
                                }
                                catch { /* skip */ }
                            }
                        }
                        catch (Exception ex) { Log($"Search for first item warning: {ex.Message}"); }

                        if (firstItem != null)
                        {
                            string itemName = SafeGet(() => firstItem.Properties.Name.Value ?? "");
                            Log($"Invoking first item: Name='{itemName}'...");
                            firstItem.Patterns.Invoke.Pattern.Invoke();
                            Log("First item invoked.");
                        }
                        else
                        {
                            Log("No Invoke-capable item found in popup to click.");
                        }
                    }
                }

                Log("--- Context Menu Explorer completed ---");
                if (outputFile != null && fileWriter != null)
                    Log($"Report saved to: {outputFile}");
            }
            finally
            {
                Console.SetOut(originalOut);
                fileWriter?.Dispose();
            }

            Console.WriteLine();
            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }

        // ─── Helpers ──────────────────────────────────────────────────────────────────

        static string SafeGet(Func<string> fn, string fallback = "")
        {
            try { return fn(); } catch { return fallback; }
        }

        static T SafeGet<T>(Func<T> fn, T fallback)
        {
            try { return fn(); } catch { return fallback; }
        }

        // ─── TeeWriter ────────────────────────────────────────────────────────────────

        /// <summary>
        /// A <see cref="TextWriter"/> that mirrors every write to two underlying writers
        /// simultaneously (e.g. the original Console.Out and a file stream).
        /// This guarantees that the file and the console receive byte-for-byte identical output.
        /// </summary>
        sealed class TeeWriter : TextWriter
        {
            private readonly TextWriter _primary;
            private readonly TextWriter _secondary;

            public TeeWriter(TextWriter primary, TextWriter secondary)
            {
                _primary   = primary;
                _secondary = secondary;
            }

            public override System.Text.Encoding Encoding => _primary.Encoding;

            public override void Write(char value)
            {
                _primary.Write(value);
                _secondary.Write(value);
            }

            public override void Write(char[] buffer, int index, int count)
            {
                _primary.Write(buffer, index, count);
                _secondary.Write(buffer, index, count);
            }

            public override void Write(string? value)
            {
                _primary.Write(value);
                _secondary.Write(value);
            }

            public override void WriteLine(string? value)
            {
                _primary.WriteLine(value);
                _secondary.WriteLine(value);
            }

            public override void Flush()
            {
                _primary.Flush();
                _secondary.Flush();
            }

            protected override void Dispose(bool disposing)
            {
                if (disposing)
                {
                    _primary.Flush();
                    _secondary.Flush();
                    // Only dispose secondary (the file); primary is Console.Out — do not close it.
                    _secondary.Dispose();
                }
                base.Dispose(disposing);
            }
        }
    }
}
