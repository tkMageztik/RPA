using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
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
        // Win32 P/Invoke — needed to give the main app window keyboard focus before sending Enter.
        [DllImport("user32.dll")] static extern bool SetForegroundWindow(IntPtr hWnd);

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
            TextWriter originalOut = Console.Out;
            StreamWriter? fileWriter = null;
            if (outputFile != null)
            {
                try { fileWriter = new StreamWriter(outputFile, false, Encoding.UTF8) { AutoFlush = true }; Console.SetOut(new TeeWriter(originalOut, fileWriter)); }
                catch (Exception ex) { LogError("Could not open output file — console only.", ex); }
            }

            try
            {
                Log("═══════════════════════════════════════════════════════════════");
                Log("  --dots: find '...' buttons, detect context menu, click item  ");
                Log("═══════════════════════════════════════════════════════════════");

                // ── Step 1: Discover '...' button candidates ─────────────────────────
                Log("[STEP-1] Discovering '...' button candidates...");
                var candidates = new List<(AutomationElement el, string label, System.Drawing.Point center)>();
                foreach (var desc in mainWindow.FindAllDescendants())
                {
                    try
                    {
                        if (!desc.Patterns.Invoke.IsSupported) continue;
                        if (SafeGet(() => desc.Properties.IsOffscreen.Value, true)) continue;
                        if (!SafeGet(() => desc.Properties.IsEnabled.Value, false)) continue;
                        string name  = SafeGet(() => desc.Properties.Name.Value ?? "");
                        string autId = SafeGet(() => desc.Properties.AutomationId.Value ?? "");
                        var    ct    = SafeGet(() => desc.Properties.ControlType.Value, ControlType.Unknown);
                        var    rect  = SafeGet(() => desc.Properties.BoundingRectangle.Value, new System.Drawing.Rectangle());
                        bool match =
                            (string.IsNullOrEmpty(name) && ct == ControlType.Button && rect.Width <= 64 && rect.Height <= 64 && !rect.IsEmpty) ||
                            name == "..." || name == "…" ||
                            name.IndexOf("more",     StringComparison.OrdinalIgnoreCase) >= 0 ||
                            name.IndexOf("option",   StringComparison.OrdinalIgnoreCase) >= 0 ||
                            name.IndexOf("menu",     StringComparison.OrdinalIgnoreCase) >= 0 ||
                            autId.IndexOf("more",     StringComparison.OrdinalIgnoreCase) >= 0 ||
                            autId.IndexOf("overflow", StringComparison.OrdinalIgnoreCase) >= 0 ||
                            autId.IndexOf("dots",     StringComparison.OrdinalIgnoreCase) >= 0 ||
                            autId.IndexOf("ellipsis", StringComparison.OrdinalIgnoreCase) >= 0;
                        if (!match) continue;
                        int cx = rect.X + rect.Width / 2, cy = rect.Y + rect.Height / 2;
                        string label = $"[{ct}] Name=\"{name}\" AutoId=\"{autId}\" X={rect.X} Y={rect.Y} W={rect.Width} H={rect.Height} center=({cx},{cy})";
                        candidates.Add((desc, label, new System.Drawing.Point(cx, cy)));
                    }
                    catch { }
                }
                if (candidates.Count == 0) { Log("No '...' candidates found. Run --poc to inspect the full element tree."); return; }
                Log($"Found {candidates.Count} candidate(s):");
                for (int i = 0; i < candidates.Count; i++) Log($"  [{i + 1}] {candidates[i].label}");

                // ── Step 2: Snapshot WinUI PopupWindowSiteBridge (if present) ─────────
                AutomationElement? popupBridge = null;
                var preClickBridgeChildIds = new HashSet<string>();
                foreach (var desc in mainWindow.FindAllDescendants())
                {
                    try
                    {
                        string cls = SafeGet(() => desc.Properties.ClassName.Value ?? "");
                        if (!cls.Contains("PopupWindowSiteBridge") && !cls.Contains("PopupHost")) continue;
                        popupBridge = desc;
                        foreach (var c in desc.FindAllChildren())
                            try { preClickBridgeChildIds.Add(c.Properties.RuntimeId.Value.ToString()); } catch { }
                        break;
                    }
                    catch { }
                }
                Log(popupBridge != null ? "[STEP-2] PopupWindowSiteBridge found." : "[STEP-2] No PopupWindowSiteBridge — standard strategies only.");

                // ── Step 3: Try each candidate ────────────────────────────────────────
                for (int ci = 0; ci < candidates.Count; ci++)
                {
                    var (candEl, candLabel, candCenter) = candidates[ci];
                    Log($"─────────────────────────────────────────────────────────");
                    Log($"[CANDIDATE {ci + 1}/{candidates.Count}] {candLabel}");

                    // Snapshot desktop windows + main-window Invoke/Menu elements before click
                    var snapWindowIds = new HashSet<string>();
                    var snapDescIds   = new HashSet<string>();
                    foreach (var w in automation.GetDesktop().FindAllChildren())
                        try { snapWindowIds.Add(w.Properties.RuntimeId.Value.ToString()); } catch { }
                    foreach (var d in mainWindow.FindAllDescendants())
                        try { if (d.Patterns.Invoke.IsSupported || d.Properties.ControlType.Value == ControlType.Menu || d.Properties.ControlType.Value == ControlType.MenuItem) snapDescIds.Add(d.Properties.RuntimeId.Value.ToString()); } catch { }
                    try { mainWindow.Focus(); } catch { }

                    // [DOTS-CLICK] Click the '...' button
                    Log($"[DOTS-CLICK] Invoking '...' button...");
                    try
                    {
                        candEl.Patterns.Invoke.Pattern.Invoke();
                        Log("[DOTS-CLICK] ✓ InvokePattern.Invoke() succeeded.");
                    }
                    catch (Exception ex)
                    {
                        Log($"[DOTS-CLICK] InvokePattern failed ({ex.Message}) — trying Mouse.Click at ({candCenter.X},{candCenter.Y})...");
                        try { Mouse.Click(candCenter); Log("[DOTS-CLICK] ✓ Mouse.Click succeeded."); }
                        catch (Exception mex) { Log($"[DOTS-CLICK] ✗ Both failed: {mex.Message}. Skipping."); continue; }
                    }

                    // ── Menu detection: S0 → S1 → S2 → S3 ───────────────────────────
                    AutomationElement? menuRoot = null;
                    string usedStrategy = "none";

                    // S0: WinUI popup (polls 3×200ms)
                    Log("[S0] WinUI popup — polling 3×200ms...");
                    for (int poll = 1; poll <= 3 && menuRoot == null; poll++)
                    {
                        System.Threading.Thread.Sleep(200);
                        try
                        {
                            // A: PopupWindowSiteBridge became visible with new children
                            AutomationElement? bridge = popupBridge;
                            if (bridge == null)
                                foreach (var d in mainWindow.FindAllDescendants())
                                { try { string c2 = SafeGet(() => d.Properties.ClassName.Value ?? ""); if (c2.Contains("PopupWindowSiteBridge") || c2.Contains("PopupHost")) { bridge = d; break; } } catch { } }
                            if (bridge != null)
                            {
                                bool vis = !SafeGet(() => bridge.Properties.IsOffscreen.Value, true);
                                var newKids = new List<AutomationElement>();
                                foreach (var bc in bridge.FindAllChildren()) try { if (!preClickBridgeChildIds.Contains(bc.Properties.RuntimeId.Value.ToString())) newKids.Add(bc); } catch { }
                                if (vis && newKids.Count > 0)
                                { menuRoot = bridge; usedStrategy = "S0-Bridge"; Log($"[S0] ✓ Bridge visible with {newKids.Count} new child(ren) (poll {poll})."); break; }
                                else Log($"[S0]   poll {poll}: visible={vis} newChildren={newKids.Count}");
                            }
                            // B: Xaml_WindowedPopupClass appeared as new desktop window
                            foreach (var dw in automation.GetDesktop().FindAllChildren())
                            {
                                try
                                {
                                    if (snapWindowIds.Contains(dw.Properties.RuntimeId.Value.ToString())) continue;
                                    string cls2 = SafeGet(() => dw.Properties.ClassName.Value ?? "");
                                    if (cls2.Contains("Xaml_WindowedPopupClass") || cls2.Contains("PopupWindowSiteBridge"))
                                    { menuRoot = dw; usedStrategy = "S0-XamlPopup"; Log($"[S0] ✓ {cls2} appeared (poll {poll})."); break; }
                                }
                                catch { }
                                if (menuRoot != null) break;
                            }
                        }
                        catch (Exception ex) { Log($"[S0] ERROR poll {poll}: {ex.Message}"); }
                    }
                    if (menuRoot == null) Log("[S0] ✗ NOT-FOUND.");

                    // S1: new top-level window (200ms)
                    if (menuRoot != null) { Log("[S1] SKIPPED — S0 found menu."); }
                    else
                    {
                        Log("[S1] New top-level window — 200ms...");
                        System.Threading.Thread.Sleep(200);
                        foreach (var w in automation.GetDesktop().FindAllChildren())
                        {
                            try
                            {
                                if (snapWindowIds.Contains(w.Properties.RuntimeId.Value.ToString())) continue;
                                string t = SafeGet(() => w.AsWindow()?.Title ?? ""), c = SafeGet(() => w.Properties.ClassName.Value ?? "");
                                Log($"[S1] ✓ New window Title='{t}' Class='{c}'");
                                if (menuRoot == null) { menuRoot = w; usedStrategy = "S1-Window"; }
                            }
                            catch { }
                        }
                        if (menuRoot == null) Log("[S1] ✗ NOT-FOUND.");
                    }

                    // S2: new Menu/MenuItem descendants (200ms)
                    if (menuRoot != null && usedStrategy != "S1-Window") { Log("[S2] SKIPPED."); }
                    else
                    {
                        Log("[S2] New Menu/MenuItem descendants — 200ms...");
                        System.Threading.Thread.Sleep(200);
                        AutomationElement? s2Root = null; int s2Count = 0;
                        foreach (var d in mainWindow.FindAllDescendants())
                        {
                            try
                            {
                                if (snapDescIds.Contains(d.Properties.RuntimeId.Value.ToString())) continue;
                                var ct = SafeGet(() => d.Properties.ControlType.Value, ControlType.Unknown);
                                if (ct != ControlType.Menu && ct != ControlType.MenuItem) continue;
                                string n = SafeGet(() => d.Properties.Name.Value ?? ""); var rr = SafeGet(() => d.Properties.BoundingRectangle.Value, new System.Drawing.Rectangle());
                                Log($"[S2]   [{ct}] Name=\"{n}\" X={rr.X} Y={rr.Y}");
                                if (ct == ControlType.Menu && s2Root == null) s2Root = d;
                                s2Count++;
                            }
                            catch { }
                        }
                        if (s2Count > 0) { menuRoot = s2Root ?? mainWindow; usedStrategy = "S2-MenuItem"; Log($"[S2] ✓ {s2Count} Menu/MenuItem found."); }
                        else Log("[S2] ✗ NOT-FOUND.");
                    }

                    // S3: new Invoke-capable descendants (200ms)
                    if (menuRoot != null && usedStrategy != "S1-Window" && usedStrategy != "S2-MenuItem") { Log("[S3] SKIPPED."); }
                    else
                    {
                        Log("[S3] New Invoke-capable descendants — 200ms...");
                        System.Threading.Thread.Sleep(200);
                        int s3Count = 0;
                        foreach (var d in mainWindow.FindAllDescendants())
                        {
                            try
                            {
                                if (snapDescIds.Contains(d.Properties.RuntimeId.Value.ToString())) continue;
                                if (!d.Patterns.Invoke.IsSupported || d.Properties.IsOffscreen.Value || !d.Properties.IsEnabled.Value) continue;
                                string n = SafeGet(() => d.Properties.Name.Value ?? ""); var rr = SafeGet(() => d.Properties.BoundingRectangle.Value, new System.Drawing.Rectangle());
                                Log($"[S3]   Name=\"{n}\" X={rr.X} Y={rr.Y}");
                                s3Count++;
                            }
                            catch { }
                        }
                        if (s3Count > 0) { menuRoot = mainWindow; usedStrategy = "S3-Invoke"; Log($"[S3] ✓ {s3Count} Invoke-capable found."); }
                        else Log("[S3] ✗ NOT-FOUND.");
                    }

                    if (menuRoot == null) { Log($"[CANDIDATE {ci + 1}] No menu found. Skipping."); continue; }

                    // ── Step 4: Find menu items and activate the first one ────────────
                    Log($"[STEP-4] Menu found via {usedStrategy}. Finding first item...");
                    var menuItems = new List<(AutomationElement el, string name, System.Drawing.Point center)>();
                    foreach (var d in menuRoot.FindAllDescendants())
                    {
                        try
                        {
                            if (snapDescIds.Contains(d.Properties.RuntimeId.Value.ToString())) continue;
                            if (!d.Patterns.Invoke.IsSupported || d.Properties.IsOffscreen.Value || !d.Properties.IsEnabled.Value) continue;
                            string n = SafeGet(() => d.Properties.Name.Value ?? ""); var rr = SafeGet(() => d.Properties.BoundingRectangle.Value, new System.Drawing.Rectangle());
                            int cx = rr.X + rr.Width / 2, cy = rr.Y + rr.Height / 2;
                            Log($"[ITEM] Name=\"{n}\" center=({cx},{cy})");
                            menuItems.Add((d, n, new System.Drawing.Point(cx, cy)));
                        }
                        catch { }
                    }

                    if (menuItems.Count == 0)
                    {
                        Log("[ACTIVATE] No Invoke items — last-resort mouse click on first visible child...");
                        foreach (var child in menuRoot.FindAllChildren())
                        {
                            try
                            {
                                var rr = SafeGet(() => child.Properties.BoundingRectangle.Value, new System.Drawing.Rectangle());
                                if (rr.IsEmpty || SafeGet(() => child.Properties.IsOffscreen.Value, true)) continue;
                                int cx = rr.X + rr.Width / 2, cy = rr.Y + rr.Height / 2;
                                Mouse.Click(new System.Drawing.Point(cx, cy));
                                Log($"[ACTIVATE] ✓ Mouse.Click at ({cx},{cy}).");
                                break;
                            }
                            catch (Exception ex) { Log($"[ACTIVATE] ✗ {ex.Message}"); }
                        }
                    }
                    else
                    {
                        var (firstEl, firstName, _) = menuItems[0];
                        Log($"[ACTIVATE] First item: \"{firstName}\". Focusing item, then sending Enter to main app window...");
                        // Select/highlight the item via InvokePattern (same as pressing an arrow key)
                        try { firstEl.Patterns.Invoke.Pattern.Invoke(); } catch { }
                        // The context menu is inline — return keyboard focus to the main app window
                        // so the subsequent Enter key-press reaches it (not our console window).
                        try
                        {
                            IntPtr mainHwnd = (IntPtr)mainWindow.Properties.NativeWindowHandle.Value;
                            if (mainHwnd != IntPtr.Zero) { SetForegroundWindow(mainHwnd); Log("[ACTIVATE] SetForegroundWindow(mainWindow) called."); }
                        }
                        catch (Exception ex) { Log($"[ACTIVATE] SetForegroundWindow warning: {ex.Message}"); }
                        System.Threading.Thread.Sleep(80);
                        Keyboard.Press(FlaUI.Core.WindowsAPI.VirtualKeyShort.RETURN);
                        Keyboard.Release(FlaUI.Core.WindowsAPI.VirtualKeyShort.RETURN);
                        Log("[ACTIVATE] ✓ Enter key sent to main app window.");
                    }

                    // ── Step 5: Wait for result window and dump it ────────────────────
                    Log("[RESULT] Waiting for result window (up to 5s)...");
                    System.Threading.Thread.Sleep(300);
                    var retryResult = Retry.WhileNull(() =>
                    {
                        foreach (var w in automation.GetDesktop().FindAllChildren())
                        {
                            try
                            {
                                if (snapWindowIds.Contains(w.Properties.RuntimeId.Value.ToString())) continue;
                                string t = SafeGet(() => w.AsWindow()?.Title ?? ""), c = SafeGet(() => w.Properties.ClassName.Value ?? "");
                                var rr = SafeGet(() => w.Properties.BoundingRectangle.Value, new System.Drawing.Rectangle());
                                if (rr.Width >= 200 && rr.Height >= 100 || !string.IsNullOrEmpty(t))
                                { Log($"[RESULT] New window Title='{t}' Class='{c}' Size={rr.Width}x{rr.Height}"); return w; }
                            }
                            catch { }
                        }
                        return null;
                    }, TimeSpan.FromSeconds(5), throwOnTimeout: false);

                    AutomationElement? resultWindow = retryResult?.Result;
                    if (resultWindow != null)
                    {
                        Log("[RESULT] ✓ Dumping result window tree...");
                        var sb = new StringBuilder();
                        sb.AppendLine("╔══════════════════════════════════════════════════╗");
                        sb.AppendLine("║   Result Window (after first menu item clicked)  ║");
                        sb.AppendLine("╚══════════════════════════════════════════════════╝");
                        int nodeCount = PrintRichUiaTree(resultWindow, sb, 0, 8);
                        Console.WriteLine(sb);
                        Log($"[RESULT] ✓ {nodeCount} element(s) dumped.");
                    }
                    else
                    {
                        Log("[RESULT] ✗ No new window within 5s. Run --poc to see what changed in the main window.");
                    }
                }

                Log("═══════════════════════════════════════════════════════════════");
                Log("  --dots COMPLETE");
                Log("═══════════════════════════════════════════════════════════════");
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
