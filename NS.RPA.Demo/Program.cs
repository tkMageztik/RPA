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
                bool dotsMode = false;
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
                    else if (args[i].Equals("--dots", StringComparison.OrdinalIgnoreCase))
                        dotsMode = true;
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

                string mode = dotsMode ? "Dots Auto-Explorer" :
                              contextMode ? "Context Menu Explorer" :
                              pocMode ? "POC Inspector" : "Demo (keyboard)";
                Log($"Mode        : {mode}");
                Log($"Application : {appPath}");
                if (windowTitle  != null) Log($"Window title: '{windowTitle}' (override)");
                if (outputFile   != null) Log($"Output file : {outputFile}");
                if (clickByName  != null) Log($"Click name  : '{clickByName}'");
                if (clickById    != null) Log($"Click id    : '{clickById}'");
                if (clickAtCoords != null) Log($"Click at    : {clickAtCoords}");

                using var automation = new UIA3Automation();

                var (_, mainWindow) = GetOrLaunchApp(appPath, automation, windowTitle);

                if (dotsMode)
                    RunDotsExplorer(mainWindow, automation, outputFile);
                else if (contextMode)
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
                if (className.Contains("WinUIDesktopWin32WindowClass") || className.Contains("WinUI"))
                    return "WinUI / Windows App SDK";
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
            Keyboard.Release(FlaUI.Core.WindowsAPI.VirtualKeyShort.RETURN);
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

        // ─── Dots Auto-Explorer mode ─────────────────────────────────────────────────

        /// <summary>
        /// Automatically discovers every "..." / more-options button in the window,
        /// clicks each one in turn, detects whatever menu or flyout appears using four
        /// strategies, clicks the first menu item, then captures and dumps the resulting
        /// new window — all without requiring any parameters from the user.
        /// </summary>
        static void RunDotsExplorer(Window mainWindow, UIA3Automation automation, string? outputFile)
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
                catch (Exception ex) { LogError($"Could not open output file — console only.", ex); }
            }

            try
            {
                Log("═══════════════════════════════════════════════════════════════════════");
                Log("  --dots Auto-Explorer: finds '...' buttons and explores their menus  ");
                Log("═══════════════════════════════════════════════════════════════════════");

                // ── Step 1: Find all "..." button candidates ──────────────────────────
                Log("Step 1: Discovering '...' / more-options button candidates...");

                var candidates = new List<(AutomationElement el, string label, System.Drawing.Point center)>();
                try
                {
                    foreach (var desc in mainWindow.FindAllDescendants())
                    {
                        try
                        {
                            // Only consider elements that are Invoke-capable and on-screen.
                            if (!desc.Patterns.Invoke.IsSupported) continue;
                            bool offscreen = SafeGet(() => desc.Properties.IsOffscreen.Value, true);
                            if (offscreen) continue;
                            bool enabled = SafeGet(() => desc.Properties.IsEnabled.Value, false);
                            if (!enabled) continue;

                            string name  = SafeGet(() => desc.Properties.Name.Value ?? "");
                            string cls   = SafeGet(() => desc.Properties.ClassName.Value ?? "");
                            string autId = SafeGet(() => desc.Properties.AutomationId.Value ?? "");
                            var    ct    = SafeGet(() => desc.Properties.ControlType.Value, ControlType.Unknown);
                            var    rect  = SafeGet(() => desc.Properties.BoundingRectangle.Value, new System.Drawing.Rectangle());

                            // Heuristic 1: empty name + Button class + small (≤64px) — the classic "..." icon button
                            bool emptySmallButton = string.IsNullOrEmpty(name)
                                && ct == ControlType.Button
                                && rect.Width <= 64 && rect.Height <= 64
                                && !rect.IsEmpty;

                            // Heuristic 2: name is literally "..." or "…"
                            bool dotName = name == "..." || name == "…";

                            // Heuristic 3: name or AutomationId contains keywords
                            bool keywordMatch =
                                name.IndexOf("more", StringComparison.OrdinalIgnoreCase)    >= 0 ||
                                name.IndexOf("option", StringComparison.OrdinalIgnoreCase)  >= 0 ||
                                name.IndexOf("menu", StringComparison.OrdinalIgnoreCase)    >= 0 ||
                                autId.IndexOf("more", StringComparison.OrdinalIgnoreCase)   >= 0 ||
                                autId.IndexOf("overflow", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                autId.IndexOf("dots", StringComparison.OrdinalIgnoreCase)   >= 0 ||
                                autId.IndexOf("ellipsis", StringComparison.OrdinalIgnoreCase) >= 0;

                            if (!emptySmallButton && !dotName && !keywordMatch) continue;

                            int cx = rect.X + rect.Width  / 2;
                            int cy = rect.Y + rect.Height / 2;
                            string reason = emptySmallButton ? "empty-name+small-button"
                                          : dotName          ? "name='...'"
                                          :                    "keyword-match";
                            string label = $"[{ct}] Name=\"{name}\" AutomationId=\"{autId}\" " +
                                           $"BoundingRect=X={rect.X} Y={rect.Y} W={rect.Width} H={rect.Height} " +
                                           $"(center {cx},{cy}) reason={reason}";

                            candidates.Add((desc, label, new System.Drawing.Point(cx, cy)));
                        }
                        catch { /* skip inaccessible */ }
                    }
                }
                catch (Exception ex) { Log($"Candidate discovery warning: {ex.Message}"); }

                if (candidates.Count == 0)
                {
                    Log("No '...' button candidates found. Run --poc to inspect the full element tree.");
                    return;
                }

                Log($"Found {candidates.Count} candidate(s):");
                for (int i = 0; i < candidates.Count; i++)
                    Log($"  [{i + 1}] {candidates[i].label}");

                // ── Step 2: Also find the WinUI PopupWindowSiteBridge pane ─────────────
                // WinUI flyouts render inside a Microsoft.UI.Content.PopupWindowSiteBridge
                // pane that is Offscreen=True when no popup is open. When a flyout opens,
                // the pane becomes visible and gains children. We snapshot its children
                // before clicking so we can diff after.
                AutomationElement? popupBridge = null;
                var preClickBridgeChildIds = new HashSet<string>();
                try
                {
                    foreach (var desc in mainWindow.FindAllDescendants())
                    {
                        try
                        {
                            string cls = SafeGet(() => desc.Properties.ClassName.Value ?? "");
                            if (cls.Contains("PopupWindowSiteBridge") || cls.Contains("PopupHost"))
                            {
                                popupBridge = desc;
                                // snapshot its current children IDs (likely empty while offscreen)
                                foreach (var child in desc.FindAllChildren())
                                    try { preClickBridgeChildIds.Add(child.Properties.RuntimeId.Value.ToString()); }
                                    catch { }
                                break;
                            }
                        }
                        catch { }
                    }
                }
                catch { }

                if (popupBridge != null)
                    Log($"WinUI PopupWindowSiteBridge found — will check for flyout content after each click.");
                else
                    Log("WinUI PopupWindowSiteBridge not found — will use standard strategies only.");

                // Summary of results across all candidates
                var successSummary = new List<string>();

                // ── Step 3: Try each candidate ────────────────────────────────────────
                for (int ci = 0; ci < candidates.Count; ci++)
                {
                    var (candEl, candLabel, candCenter) = candidates[ci];
                    Log("───────────────────────────────────────────────────────────────────");
                    Log($"Trying candidate [{ci + 1}/{candidates.Count}]: {candLabel}");

                    // Per-candidate decision record for the final table
                    string r_click    = "?";   // how the "..." button was clicked
                    string r_s0       = "?";   // S0 result
                    string r_s1       = "?";   // S1 result
                    string r_s2       = "?";   // S2 result
                    string r_s3       = "?";   // S3 result
                    string r_activate = "?";   // how first item was activated
                    string r_window   = "?";   // result window

                    // Take fresh snapshot of top-level windows and main window descendants
                    var snapWindowIds = new HashSet<string>();
                    var snapDescIds   = new HashSet<string>();
                    try
                    {
                        foreach (var w in automation.GetDesktop().FindAllChildren())
                            try { snapWindowIds.Add(w.Properties.RuntimeId.Value.ToString()); } catch { }
                        foreach (var d in mainWindow.FindAllDescendants())
                            try { if (d.Patterns.Invoke.IsSupported || d.Properties.ControlType.Value == ControlType.Menu || d.Properties.ControlType.Value == ControlType.MenuItem)
                                    snapDescIds.Add(d.Properties.RuntimeId.Value.ToString()); } catch { }
                    }
                    catch { }

                    // Ensure the window is focused before clicking
                    try { mainWindow.Focus(); } catch { }

                    // Click the candidate
                    try
                    {
                        Log($"  [CLICK] Sending InvokePattern to '...' button...");
                        candEl.Patterns.Invoke.Pattern.Invoke();
                        r_click = "✓ InvokePattern";
                        Log($"  [CLICK] ✓ InvokePattern.Invoke() succeeded.");
                    }
                    catch (Exception ex)
                    {
                        Log($"  [CLICK] ✗ InvokePattern failed ({ex.Message}) — trying mouse click at ({candCenter.X},{candCenter.Y})...");
                        try
                        {
                            Mouse.Click(candCenter);
                            r_click = "✓ Mouse.Click fallback";
                            Log("  [CLICK] ✓ Mouse click succeeded.");
                        }
                        catch (Exception mex)
                        {
                            r_click = $"✗ both failed: {mex.Message}";
                            Log($"  [CLICK] ✗ Mouse click also failed: {mex.Message}. Skipping candidate.");
                            successSummary.Add($"[{ci + 1}] CLICK={r_click} — candidate skipped");
                            continue;
                        }
                    }

                    AutomationElement? menuRoot = null;
                    string usedStrategy = "none";
                    string menuSummary  = "";

                    // ── Strategy 0: WinUI PopupWindowSiteBridge / Xaml_WindowedPopupClass ──
                    // WinUI lazily adds PopupWindowSiteBridge to the tree on first flyout.
                    // We poll up to 3×200 ms. Each poll also checks desktop-level
                    // Xaml_WindowedPopupClass HWNDs (the actual WinUI 3 flyout window).
                    Log("  [S0] WinUI popup detection — polling up to 600 ms (3×200 ms)...");
                    for (int s0Poll = 1; s0Poll <= 3 && menuRoot == null; s0Poll++)
                    {
                        System.Threading.Thread.Sleep(200);
                        Log($"  [S0] poll {s0Poll}/3...");
                        try
                        {
                            // Sub-strategy A: check existing cached bridge (or re-find it)
                            AutomationElement? bridge = popupBridge;
                            if (bridge == null)
                            {
                                // Re-scan: WinUI may have lazily added it after the first click
                                foreach (var d in mainWindow.FindAllDescendants())
                                {
                                    try
                                    {
                                        string cls2 = SafeGet(() => d.Properties.ClassName.Value ?? "");
                                        if (cls2.Contains("PopupWindowSiteBridge") || cls2.Contains("PopupHost"))
                                        { bridge = d; Log($"  [S0]   PopupWindowSiteBridge appeared in tree (lazy add)."); break; }
                                    }
                                    catch { }
                                }
                            }
                            if (bridge != null)
                            {
                                bool bridgeVisible = !SafeGet(() => bridge.Properties.IsOffscreen.Value, true);
                                var  bridgeChildren = bridge.FindAllChildren();
                                var  newBridgeChildren = new List<AutomationElement>();
                                foreach (var bc in bridgeChildren)
                                    try { if (!preClickBridgeChildIds.Contains(bc.Properties.RuntimeId.Value.ToString())) newBridgeChildren.Add(bc); } catch { }
                                if (bridgeVisible && newBridgeChildren.Count > 0)
                                {
                                    r_s0 = $"✓ WORKED (A-Bridge poll {s0Poll}) — {newBridgeChildren.Count} new child(ren)";
                                    Log($"  [S0] {r_s0}");
                                    foreach (var bc in newBridgeChildren)
                                        Log($"  [S0]   Child: [{SafeGet(() => bc.Properties.ControlType.Value.ToString(), "?")}] Name=\"{SafeGet(() => bc.Properties.Name.Value ?? "")}\"");
                                    menuRoot = bridge; usedStrategy = "S0-PopupBridge";
                                    break;
                                }
                                else Log($"  [S0]   A-Bridge: visible={bridgeVisible} newChildren={newBridgeChildren.Count}");
                            }

                            // Sub-strategy B: Xaml_WindowedPopupClass desktop window (WinUI 3 flyout HWND)
                            foreach (var deskChild in automation.GetDesktop().FindAllChildren())
                            {
                                try
                                {
                                    string rid = deskChild.Properties.RuntimeId.Value.ToString();
                                    if (snapWindowIds.Contains(rid)) continue; // existed before click
                                    string cls2 = SafeGet(() => deskChild.Properties.ClassName.Value ?? "");
                                    if (cls2.Contains("Xaml_WindowedPopupClass") || cls2.Contains("PopupWindowSiteBridge") || cls2.Contains("DesktopChildSiteBridge"))
                                    {
                                        string t2 = SafeGet(() => deskChild.Properties.Name.Value ?? "");
                                        r_s0 = $"✓ WORKED (B-XamlPopup poll {s0Poll}) — Class='{cls2}' Name='{t2}'";
                                        Log($"  [S0] {r_s0}");
                                        menuRoot = deskChild; usedStrategy = "S0-XamlPopup";
                                        break;
                                    }
                                }
                                catch { }
                                if (menuRoot != null) break;
                            }
                        }
                        catch (Exception ex) { r_s0 = $"✗ ERROR poll {s0Poll}: {ex.Message}"; Log($"  [S0] {r_s0}"); }
                    }
                    if (menuRoot == null)
                    {
                        r_s0 = popupBridge == null
                            ? "✗ NOT-FOUND — PopupWindowSiteBridge never appeared + no Xaml_WindowedPopupClass"
                            : "✗ NOT-FOUND — bridge offscreen/no new children + no Xaml_WindowedPopupClass";
                        Log($"  [S0] {r_s0}");
                    }

                    // ── Strategy 1 (400 ms): new top-level window ─────────────────────
                    if (menuRoot != null)
                    {
                        r_s1 = "✗ SKIPPED — S0 already found menu";
                        Log($"  [S1] {r_s1}");
                    }
                    else
                    {
                        Log("  [S1] New top-level window — sleeping 200 ms (S0 already waited 600 ms)...");
                        System.Threading.Thread.Sleep(200);
                        try
                        {
                            bool s1found = false;
                            foreach (var child in automation.GetDesktop().FindAllChildren())
                            {
                                try
                                {
                                    string rid = child.Properties.RuntimeId.Value.ToString();
                                    if (!snapWindowIds.Contains(rid))
                                    {
                                        string t = SafeGet(() => child.AsWindow()?.Title ?? "");
                                        string c = SafeGet(() => child.Properties.ClassName.Value ?? "");
                                        Log($"  [S1]   New window Title='{t}' Class='{c}'");
                                        if (menuRoot == null) { menuRoot = child; usedStrategy = "S1-NewWindow"; }
                                        s1found = true;
                                    }
                                }
                                catch { }
                            }
                            r_s1 = s1found ? "✓ WORKED — new top-level window appeared" : "✗ NOT-FOUND — no new top-level window";
                            Log($"  [S1] {r_s1}");
                        }
                        catch (Exception ex) { r_s1 = $"✗ ERROR: {ex.Message}"; Log($"  [S1] {r_s1}"); }
                    }

                    // ── Strategy 2 (600 ms): new Menu/MenuItem descendants ─────────────
                    if (menuRoot != null && usedStrategy != "S1-NewWindow")
                    {
                        r_s2 = "✗ SKIPPED — earlier strategy already found menu";
                        Log($"  [S2] {r_s2}");
                    }
                    else
                    {
                        Log("  [S2] New Menu/MenuItem descendants — sleeping 200 ms more (600 ms total)...");
                        System.Threading.Thread.Sleep(200);
                        int s2Count = 0;
                        AutomationElement? s2Root = null;
                        try
                        {
                            foreach (var desc in mainWindow.FindAllDescendants())
                            {
                                try
                                {
                                    string rid = desc.Properties.RuntimeId.Value.ToString();
                                    if (snapDescIds.Contains(rid)) continue;
                                    var ct = SafeGet(() => desc.Properties.ControlType.Value, ControlType.Unknown);
                                    if (ct == ControlType.Menu || ct == ControlType.MenuItem)
                                    {
                                        string n = SafeGet(() => desc.Properties.Name.Value ?? "");
                                        var r = SafeGet(() => desc.Properties.BoundingRectangle.Value, new System.Drawing.Rectangle());
                                        Log($"  [S2]   [{ct}] Name=\"{n}\" BoundingRect=X={r.X} Y={r.Y} W={r.Width} H={r.Height}");
                                        if (ct == ControlType.Menu && s2Root == null) s2Root = desc;
                                        s2Count++;
                                    }
                                }
                                catch { }
                            }
                            if (s2Count > 0)
                            {
                                r_s2 = $"✓ WORKED — {s2Count} Menu/MenuItem element(s)";
                                Log($"  [S2] {r_s2}");
                                if (menuRoot == null) { menuRoot = s2Root ?? mainWindow; usedStrategy = "S2-MenuItem"; }
                            }
                            else
                            {
                                r_s2 = "✗ NOT-FOUND — no new Menu/MenuItem";
                                Log($"  [S2] {r_s2}");
                            }
                        }
                        catch (Exception ex) { r_s2 = $"✗ ERROR: {ex.Message}"; Log($"  [S2] {r_s2}"); }
                    }

                    // ── Strategy 3 (800 ms): new Invoke-capable descendants ────────────
                    if (menuRoot != null && usedStrategy != "S1-NewWindow" && usedStrategy != "S2-MenuItem")
                    {
                        r_s3 = "✗ SKIPPED — earlier strategy already found menu";
                        Log($"  [S3] {r_s3}");
                    }
                    else
                    {
                        Log("  [S3] New Invoke-capable descendants — sleeping 200 ms more (800 ms total)...");
                        System.Threading.Thread.Sleep(200);
                        int s3Count = 0;
                        try
                        {
                            foreach (var desc in mainWindow.FindAllDescendants())
                            {
                                try
                                {
                                    string rid = desc.Properties.RuntimeId.Value.ToString();
                                    if (snapDescIds.Contains(rid)) continue;
                                    if (desc.Patterns.Invoke.IsSupported && !desc.Properties.IsOffscreen.Value && desc.Properties.IsEnabled.Value)
                                    {
                                        string n  = SafeGet(() => desc.Properties.Name.Value ?? "");
                                        string ct = SafeGet(() => desc.Properties.ControlType.Value.ToString(), "?");
                                        var r = SafeGet(() => desc.Properties.BoundingRectangle.Value, new System.Drawing.Rectangle());
                                        Log($"  [S3]   [{ct}] Name=\"{n}\" BoundingRect=X={r.X} Y={r.Y} W={r.Width} H={r.Height}");
                                        s3Count++;
                                    }
                                }
                                catch { }
                            }
                            if (s3Count > 0)
                            {
                                r_s3 = $"✓ WORKED — {s3Count} new Invoke-capable element(s)";
                                Log($"  [S3] {r_s3}");
                                if (menuRoot == null) { menuRoot = mainWindow; usedStrategy = "S3-Invoke"; }
                            }
                            else
                            {
                                r_s3 = "✗ NOT-FOUND — no new Invoke-capable elements";
                                Log($"  [S3] {r_s3}");
                            }
                        }
                        catch (Exception ex) { r_s3 = $"✗ ERROR: {ex.Message}"; Log($"  [S3] {r_s3}"); }
                    }

                    if (menuRoot == null)
                    {
                        r_activate = "✗ NOT-REACHED — no menu found";
                        r_window   = "✗ NOT-REACHED";
                        Log($"  No menu detected for candidate [{ci + 1}]. Skipping to next.");
                        successSummary.Add($"[{ci+1}] CLICK={r_click} S0={r_s0} S1={r_s1} S2={r_s2} S3={r_s3} → NO MENU");
                        continue;
                    }

                    // ── Step 4: Find all menu items and click the FIRST one ───────────
                    Log($"  [MENU-FOUND via {usedStrategy}] Step 4: Finding and clicking the first menu item...");

                    var menuItems = new List<(AutomationElement el, string name, System.Drawing.Point center)>();
                    try
                    {
                        foreach (var desc in menuRoot.FindAllDescendants())
                        {
                            try
                            {
                                string rid = desc.Properties.RuntimeId.Value.ToString();
                                if (snapDescIds.Contains(rid)) continue;
                                if (desc.Patterns.Invoke.IsSupported && !desc.Properties.IsOffscreen.Value && desc.Properties.IsEnabled.Value)
                                {
                                    string n = SafeGet(() => desc.Properties.Name.Value ?? "");
                                    var r = SafeGet(() => desc.Properties.BoundingRectangle.Value, new System.Drawing.Rectangle());
                                    int ix = r.X + r.Width / 2;
                                    int iy = r.Y + r.Height / 2;
                                    Log($"  [ITEM] Name=\"{n}\" center=({ix},{iy})");
                                    menuItems.Add((desc, n, new System.Drawing.Point(ix, iy)));
                                }
                            }
                            catch { }
                        }
                    }
                    catch (Exception ex) { Log($"  Menu item search warning: {ex.Message}"); }

                    if (menuItems.Count == 0)
                    {
                        Log($"  [ACTIVATE] No Invoke-capable items — trying last-resort: mouse click on first visible child...");
                        bool lastResortOk = false;
                        try
                        {
                            foreach (var child in menuRoot.FindAllChildren())
                            {
                                var r = SafeGet(() => child.Properties.BoundingRectangle.Value, new System.Drawing.Rectangle());
                                if (!r.IsEmpty && !SafeGet(() => child.Properties.IsOffscreen.Value, true))
                                {
                                    int ix = r.X + r.Width / 2;
                                    int iy = r.Y + r.Height / 2;
                                    Log($"  [ACTIVATE] ✓ Mouse clicking first visible child at ({ix},{iy})...");
                                    Mouse.Click(new System.Drawing.Point(ix, iy));
                                    r_activate = $"✓ last-resort Mouse.Click at ({ix},{iy})";
                                    Log("  [ACTIVATE] ✓ Click sent.");
                                    lastResortOk = true;
                                    break;
                                }
                            }
                        }
                        catch (Exception ex) { Log($"  [ACTIVATE] ✗ Last-resort click failed: {ex.Message}"); }
                        if (!lastResortOk) r_activate = "✗ no items found, last-resort also failed";
                    }
                    else
                    {
                        var (firstEl, firstName, firstCenter) = menuItems[0];
                        Log($"  [ACTIVATE] Focusing first item \"{firstName}\" via InvokePattern + Enter key-down+up...");
                        // WinUI MenuFlyoutItem: Mouse.Click moves the cursor and focuses/hovers
                        // the item (same as arrow key), but Enter is what actually activates it.
                        // Strategy: focus via InvokePattern, then send full Enter key-down+key-up.
                        // Keyboard.Press sends key-down only; we must also call Release for key-up.
                        try { firstEl.Patterns.Invoke.Pattern.Invoke(); } catch { /* focus attempt — ignore */ }
                        System.Threading.Thread.Sleep(80); // brief pause so WinUI registers focus
                        Keyboard.Press(FlaUI.Core.WindowsAPI.VirtualKeyShort.RETURN);
                        Keyboard.Release(FlaUI.Core.WindowsAPI.VirtualKeyShort.RETURN);
                        r_activate = $"✓ InvokePattern(focus) + Enter key-down+up on \"{firstName}\"";
                        Log($"  [ACTIVATE] ✓ Enter key down+up sent — item activated.");
                        menuSummary = $"menu via {usedStrategy}, {menuItems.Count} item(s), activated \"{firstName}\"";
                    }

                    // ── Step 5: Wait for resulting new window and dump it ─────────────
                    Log("  [RESULT] Waiting for result window (up to 5 seconds)...");
                    System.Threading.Thread.Sleep(300);
                    AutomationElement? resultWindow = null;
                    try
                    {
                        var retryResult = Retry.WhileNull(
                            () =>
                            {
                                try
                                {
                                    foreach (var child in automation.GetDesktop().FindAllChildren())
                                    {
                                        try
                                        {
                                            string rid = child.Properties.RuntimeId.Value.ToString();
                                            if (!snapWindowIds.Contains(rid))
                                            {
                                                string t = SafeGet(() => child.AsWindow()?.Title ?? "");
                                                string c = SafeGet(() => child.Properties.ClassName.Value ?? "");
                                                // Exclude the context menu window itself (tiny, no title)
                                                var r = SafeGet(() => child.Properties.BoundingRectangle.Value, new System.Drawing.Rectangle());
                                                bool isLargeEnough = r.Width >= 200 && r.Height >= 100;
                                                if (isLargeEnough || !string.IsNullOrEmpty(t))
                                                {
                                                    Log($"  [RESULT] New window — Title='{t}' Class='{c}' Size={r.Width}x{r.Height}");
                                                    return child;
                                                }
                                            }
                                        }
                                        catch { }
                                    }
                                }
                                catch { }
                                return null;
                            },
                            TimeSpan.FromSeconds(5),
                            throwOnTimeout: false);

                        resultWindow = retryResult?.Result;
                    }
                    catch (Exception ex) { Log($"  [RESULT] Result window wait error: {ex.Message}"); }

                    if (resultWindow != null)
                    {
                        r_window = "✓ result window appeared";
                        Log($"  [RESULT] {r_window} — dumping element tree...");
                        var sb = new StringBuilder();
                        sb.AppendLine();
                        sb.AppendLine("╔══════════════════════════════════════════════════════════════════════════════╗");
                        sb.AppendLine("║           Result Window (after clicking first menu item)                    ║");
                        sb.AppendLine("╚══════════════════════════════════════════════════════════════════════════════╝");
                        sb.AppendLine($"Candidate: [{ci + 1}] {candLabel}");
                        sb.AppendLine($"Menu     : {menuSummary}");
                        sb.AppendLine($"Captured : {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                        sb.AppendLine();
                        int nodeCount = PrintRichUiaTree(resultWindow, sb, indent: 0, maxDepth: 8);
                        Console.WriteLine(sb.ToString());
                        Log($"  [RESULT] ✓ Result window dumped: {nodeCount} element(s).");
                        successSummary.Add($"[{ci+1}] CLICK={r_click} | S0={r_s0} | S1={r_s1} | S2={r_s2} | S3={r_s3} | ACTIVATE={r_activate} | WINDOW={r_window} ({nodeCount} elements)");
                    }
                    else
                    {
                        r_window = "✗ no new window within 5s (action may have changed same window)";
                        Log($"  [RESULT] {r_window}");
                        Log("  [RESULT] Tip: run --poc again to see what changed inside the main window.");
                        successSummary.Add($"[{ci+1}] CLICK={r_click} | S0={r_s0} | S1={r_s1} | S2={r_s2} | S3={r_s3} | ACTIVATE={r_activate} | WINDOW={r_window}");
                    }
                }

                // ── Final summary / decision table ───────────────────────────────────
                Log("═══════════════════════════════════════════════════════════════════════");
                Log("  --dots Auto-Explorer COMPLETE — Decision Table");
                Log("  (✓ WORKED = keep this code | ✗ SKIPPED = dead code for this app)");
                Log("───────────────────────────────────────────────────────────────────────");
                Log($"  Candidates tried: {candidates.Count}");
                foreach (var s in successSummary)
                {
                    // Print each field on its own line for readability
                    Log($"  Candidate {s}");
                }
                Log("═══════════════════════════════════════════════════════════════════════");
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

                // ── Step 3: Try ALL detection strategies with individual sleeps ──────────
                // All 4 strategies always run — each logs FOUND/NOT FOUND so you can
                // tell which one works for your application.
                Log("Step 3: Trying all context menu detection strategies...");

                AutomationElement? popup    = null;  // first successful strategy sets this
                string             usedStrategy = "none";

                // ┌──────────────────────────────────────────────────────────────────────┐
                // │ Strategy 1 — New top-level window (200 ms wait)                     │
                // │ Classic Win32 context menus and some WPF menus open their own HWND. │
                // └──────────────────────────────────────────────────────────────────────┘
                Log("─────────────────────────────────────────────────────────────────────");
                Log("[Strategy 1] New top-level window — sleeping 200 ms...");
                System.Threading.Thread.Sleep(200);
                try
                {
                    foreach (var child in automation.GetDesktop().FindAllChildren())
                    {
                        try
                        {
                            string rid = child.Properties.RuntimeId.Value.ToString();
                            if (!existingWindowIds.Contains(rid))
                            {
                                string title = SafeGet(() => child.AsWindow()?.Title ?? "");
                                string cls   = SafeGet(() => child.Properties.ClassName.Value ?? "");
                                Log($"[Strategy 1] FOUND: new window Title='{title}' Class='{cls}'");
                                if (popup == null) { popup = child; usedStrategy = "Strategy 1"; }
                            }
                        }
                        catch { /* element gone */ }
                    }
                    if (popup == null) Log("[Strategy 1] NOT FOUND: no new top-level window.");
                }
                catch (Exception ex) { Log($"[Strategy 1] Error: {ex.Message}"); }

                // ┌──────────────────────────────────────────────────────────────────────┐
                // │ Strategy 2 — Menu/MenuItem control types in main window (400 ms)    │
                // │ Right-click context menus in WinUI/XAML add Menu or MenuItem nodes  │
                // │ directly inside the existing window hierarchy.                      │
                // └──────────────────────────────────────────────────────────────────────┘
                Log("─────────────────────────────────────────────────────────────────────");
                Log("[Strategy 2] Menu/MenuItem in main window — sleeping 200 ms more (400 ms total)...");
                System.Threading.Thread.Sleep(200);
                AutomationElement? menuRoot2 = null;
                int menuCount2 = 0;
                try
                {
                    foreach (var desc in mainWindow.FindAllDescendants())
                    {
                        try
                        {
                            string rid = desc.Properties.RuntimeId.Value.ToString();
                            if (existingDescendantIds.Contains(rid)) continue;
                            var ct = SafeGet(() => desc.Properties.ControlType.Value, ControlType.Unknown);
                            if (ct == ControlType.Menu || ct == ControlType.MenuItem)
                            {
                                string n = SafeGet(() => desc.Properties.Name.Value ?? "");
                                string ctStr = ct.ToString();
                                var r = SafeGet(() => desc.Properties.BoundingRectangle.Value, new System.Drawing.Rectangle());
                                Log($"[Strategy 2] FOUND [{ctStr}] Name=\"{n}\" BoundingRect=X={r.X} Y={r.Y} W={r.Width} H={r.Height}");
                                if (ct == ControlType.Menu && menuRoot2 == null) menuRoot2 = desc;
                                menuCount2++;
                            }
                        }
                        catch { /* skip */ }
                    }
                    if (menuCount2 > 0)
                    {
                        Log($"[Strategy 2] FOUND: {menuCount2} Menu/MenuItem element(s).");
                        if (popup == null) { popup = menuRoot2 ?? mainWindow; usedStrategy = "Strategy 2"; }
                    }
                    else Log("[Strategy 2] NOT FOUND: no new Menu/MenuItem elements.");
                }
                catch (Exception ex) { Log($"[Strategy 2] Error: {ex.Message}"); }

                // ┌──────────────────────────────────────────────────────────────────────┐
                // │ Strategy 3 — New Invoke-capable visible descendants (600 ms)        │
                // │ WinUI CommandBarFlyout items and toolbar overflow buttons appear as  │
                // │ new Button/ListItem/AppBarButton elements inside the main window.    │
                // └──────────────────────────────────────────────────────────────────────┘
                Log("─────────────────────────────────────────────────────────────────────");
                Log("[Strategy 3] New Invoke-capable descendants — sleeping 200 ms more (600 ms total)...");
                System.Threading.Thread.Sleep(200);
                int invokeCount3 = 0;
                try
                {
                    foreach (var desc in mainWindow.FindAllDescendants())
                    {
                        try
                        {
                            string rid = desc.Properties.RuntimeId.Value.ToString();
                            if (existingDescendantIds.Contains(rid)) continue;
                            if (desc.Patterns.Invoke.IsSupported
                                && !desc.Properties.IsOffscreen.Value
                                && desc.Properties.IsEnabled.Value)
                            {
                                string n = SafeGet(() => desc.Properties.Name.Value ?? "");
                                string ct = SafeGet(() => desc.Properties.ControlType.Value.ToString(), "?");
                                var r = SafeGet(() => desc.Properties.BoundingRectangle.Value, new System.Drawing.Rectangle());
                                Log($"[Strategy 3] FOUND [{ct}] Name=\"{n}\" BoundingRect=X={r.X} Y={r.Y} W={r.Width} H={r.Height} " +
                                    $"FlaUI: element.Patterns.Invoke.Pattern.Invoke()");
                                invokeCount3++;
                            }
                        }
                        catch { /* skip */ }
                    }
                    if (invokeCount3 > 0)
                    {
                        Log($"[Strategy 3] FOUND: {invokeCount3} new Invoke-capable element(s).");
                        if (popup == null) { popup = mainWindow; usedStrategy = "Strategy 3"; }
                    }
                    else Log("[Strategy 3] NOT FOUND: no new Invoke-capable elements.");
                }
                catch (Exception ex) { Log($"[Strategy 3] Error: {ex.Message}"); }

                // ┌──────────────────────────────────────────────────────────────────────┐
                // │ Strategy 4 — ExpandCollapse descendants (1000 ms)                  │
                // │ Some WinUI menus and tree views use ExpandCollapsePattern instead   │
                // │ of Invoke. Clicking "..." may expand a sub-section rather than open │
                // │ a true popup.                                                       │
                // └──────────────────────────────────────────────────────────────────────┘
                Log("─────────────────────────────────────────────────────────────────────");
                Log("[Strategy 4] New ExpandCollapse descendants — sleeping 400 ms more (1000 ms total)...");
                System.Threading.Thread.Sleep(400);
                int ecCount4 = 0;
                try
                {
                    foreach (var desc in mainWindow.FindAllDescendants())
                    {
                        try
                        {
                            string rid = desc.Properties.RuntimeId.Value.ToString();
                            if (existingDescendantIds.Contains(rid)) continue;
                            if (desc.Patterns.ExpandCollapse.IsSupported
                                && !desc.Properties.IsOffscreen.Value)
                            {
                                string n  = SafeGet(() => desc.Properties.Name.Value ?? "");
                                string ct = SafeGet(() => desc.Properties.ControlType.Value.ToString(), "?");
                                string ecState = SafeGet(() => desc.Patterns.ExpandCollapse.Pattern.ExpandCollapseState.Value.ToString(), "?");
                                var r = SafeGet(() => desc.Properties.BoundingRectangle.Value, new System.Drawing.Rectangle());
                                Log($"[Strategy 4] FOUND [{ct}] Name=\"{n}\" State={ecState} BoundingRect=X={r.X} Y={r.Y} W={r.Width} H={r.Height} " +
                                    $"FlaUI: element.Patterns.ExpandCollapse.Pattern.Expand()");
                                ecCount4++;
                            }
                        }
                        catch { /* skip */ }
                    }
                    if (ecCount4 > 0)
                    {
                        Log($"[Strategy 4] FOUND: {ecCount4} new ExpandCollapse element(s).");
                        if (popup == null) { popup = mainWindow; usedStrategy = "Strategy 4"; }
                    }
                    else Log("[Strategy 4] NOT FOUND: no new ExpandCollapse elements.");
                }
                catch (Exception ex) { Log($"[Strategy 4] Error: {ex.Message}"); }

                Log("─────────────────────────────────────────────────────────────────────");
                if (popup != null)
                    Log($"Step 3 complete — using result from {usedStrategy}. Proceeding to dump.");
                else
                    Log("Step 3 complete — NO strategy found context menu items. Check click target coordinates.");

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
                            var itemRect = SafeGet(() => firstItem.Properties.BoundingRectangle.Value, new System.Drawing.Rectangle());
                            int itemCx = itemRect.X + itemRect.Width  / 2;
                            int itemCy = itemRect.Y + itemRect.Height / 2;
                            Log($"Clicking first item: Name='{itemName}' at ({itemCx},{itemCy})...");
                            // Physical mouse click — WinUI MenuFlyoutItem requires this
                            try { firstItem.Patterns.Invoke.Pattern.Invoke(); } catch { /* focus attempt — ignore */ }
                            System.Threading.Thread.Sleep(80);
                            Keyboard.Press(FlaUI.Core.WindowsAPI.VirtualKeyShort.RETURN);
                            Keyboard.Release(FlaUI.Core.WindowsAPI.VirtualKeyShort.RETURN);
                            Log("First item activated via Enter key down+up.");
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
