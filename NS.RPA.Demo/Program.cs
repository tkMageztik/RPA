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
                string? outputFile = null;
                string? windowTitle = null;

                for (int i = 0; i < args.Length; i++)
                {
                    if (args[i].Equals("--poc", StringComparison.OrdinalIgnoreCase))
                        pocMode = true;
                    else if (args[i].Equals("--out", StringComparison.OrdinalIgnoreCase) && i + 1 < args.Length)
                        outputFile = args[++i];
                    else if (args[i].Equals("--title", StringComparison.OrdinalIgnoreCase) && i + 1 < args.Length)
                        windowTitle = args[++i];
                    else
                        appPath = args[i];
                }

                if (string.IsNullOrWhiteSpace(appPath))
                    throw new ArgumentException("Application path must not be empty.");

                Log($"Mode        : {(pocMode ? "POC Inspector" : "Demo (keyboard)")}");
                Log($"Application : {appPath}");
                if (windowTitle != null) Log($"Window title: '{windowTitle}' (override)");
                if (outputFile != null)  Log($"Output file : {outputFile}");

                using var automation = new UIA3Automation();

                var (_, mainWindow) = GetOrLaunchApp(appPath, automation, windowTitle);

                if (pocMode)
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
