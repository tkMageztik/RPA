using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Input;
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

                foreach (var arg in args)
                {
                    if (arg.Equals("--poc", StringComparison.OrdinalIgnoreCase))
                        pocMode = true;
                    else
                        appPath = arg;
                }

                if (string.IsNullOrWhiteSpace(appPath))
                    throw new ArgumentException("Application path must not be empty.");

                Log($"Mode        : {(pocMode ? "POC Inspector" : "Demo (keyboard)")}");
                Log($"Application : {appPath}");

                using var automation = new UIA3Automation();

                var (_, mainWindow) = GetOrLaunchApp(appPath, automation);

                if (pocMode)
                    RunPocInspector(mainWindow);
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
        /// Uses a multi-strategy window search that handles Win32, PowerBuilder, and UWP apps.
        /// </summary>
        /// <param name="appPath">Full path or bare executable name (e.g. "C:\Apps\MyApp.exe" or "calc.exe").</param>
        /// <param name="automation">The UIA3 automation instance used to locate windows.</param>
        /// <returns>A tuple containing the attached <see cref="Application"/> and its ready <see cref="Window"/> main window.</returns>
        static (Application app, Window mainWindow) GetOrLaunchApp(string appPath, UIA3Automation automation)
        {
            string processName = Path.GetFileNameWithoutExtension(appPath);
            Log($"Searching for existing window of '{processName}' using multi-strategy approach...");

            var existingWindow = FindWindowMultiStrategy(processName, automation);
            if (existingWindow != null)
            {
                Log($"Application is already running — attaching to existing window.");
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

            // Wait reactively using the same multi-strategy search — this handles UWP apps
            // (where calc.exe exits immediately and CalculatorApp.exe appears later) and any
            // other case where the process name differs from the executable used to launch.
            Log($"Waiting for window of '{processName}' to appear (multi-strategy)...");
            var retryResult = Retry.WhileNull(
                () => FindWindowMultiStrategy(processName, automation),
                TimeSpan.FromSeconds(30),
                throwOnTimeout: true,
                timeoutMessage: $"Window for '{processName}' did not appear within 30 seconds.");

            // throwOnTimeout:true guarantees Result is non-null here.
            var mainWindow = retryResult.Result!;
            var launchedProcess = Process.GetProcessById((int)mainWindow.Properties.ProcessId.Value);
            Log($"Window is ready — PID {launchedProcess.Id}, ProcessName '{launchedProcess.ProcessName}'.");

            return (Application.Attach(launchedProcess), mainWindow);
        }

        /// <summary>
        /// Finds the main window of a target application using three strategies in order.
        /// Returns null if no matching window is found on this attempt (caller should retry).
        /// </summary>
        /// <param name="processName">Executable name without extension (e.g. "calc", "MyPBApp").</param>
        /// <param name="automation">The UIA3 automation instance.</param>
        /// <returns>The matching <see cref="Window"/>, or null if not found.</returns>
        static Window? FindWindowMultiStrategy(string processName, UIA3Automation automation)
        {
            // ── Strategy 1: Exact process name ─────────────────────────────────────────
            // Works for: Win32 apps, PowerBuilder, classic .NET WinForms/WPF, and
            // UWP apps that have their own process (e.g. CalculatorApp).
            var procs = Process.GetProcessesByName(processName);
            foreach (var proc in procs)
            {
                try
                {
                    var app = Application.Attach(proc);
                    var win = app.GetMainWindow(automation);
                    if (win != null)
                    {
                        Log($"[Strategy 1 - Exact process] Found '{processName}' (PID {proc.Id}).");
                        DetectAndLogAppType(win);
                        return win;
                    }
                }
                catch (Exception ex) { Log($"[Strategy 1] Skipping PID {proc.Id}: {ex.Message}"); }
            }

            // ── Strategy 2: UWP ApplicationFrameHost ──────────────────────────────────
            // UWP apps (like Calculator on Windows 10/11) are visually hosted inside
            // ApplicationFrameHost.exe. The window title or ClassName often contains
            // the original process name as a hint.
            var hostProcs = Process.GetProcessesByName(UwpHostProcessName);
            foreach (var hostProc in hostProcs)
            {
                try
                {
                    var hostApp = Application.Attach(hostProc);
                    foreach (var win in hostApp.GetAllTopLevelWindows(automation))
                    {
                        string title = win.Title ?? "";
                        string className = win.Properties.ClassName.IsSupported
                            ? (win.Properties.ClassName.Value ?? "")
                            : "";

                        // Title substring match is intentional: "calc" matches "Calculator".
                        // Strategy 3 does exact process-name matching as a more precise fallback.
                        if (title.IndexOf(processName, StringComparison.OrdinalIgnoreCase) >= 0
                            || className.IndexOf(processName, StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            Log($"[Strategy 2 - UWP host] Found in '{UwpHostProcessName}' — Title: '{title}', Class: '{className}'.");
                            DetectAndLogAppType(win);
                            return win;
                        }
                    }
                }
                catch (Exception ex) { Log($"[Strategy 2] Skipping UWP host PID {hostProc.Id}: {ex.Message}"); }
            }

            // ── Strategy 3: Full desktop scan ─────────────────────────────────────────
            // Last resort: enumerate ALL top-level windows on the desktop and look for
            // one whose owning process name matches. Covers edge cases such as shell-
            // launched apps, wrapper executables, or apps with unusual process names.
            // Terminates as soon as the first match is found (early return inside loop).
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

                        if (proc.ProcessName.Equals(processName, StringComparison.OrdinalIgnoreCase))
                        {
                            Log($"[Strategy 3 - Desktop scan] Found window owned by process '{proc.ProcessName}' (PID {pid}).");
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

        /// <summary>
        /// Reads the window ClassName to detect and log the type of application framework.
        /// Helps identify UWP, PowerBuilder, WinForms, WPF, Electron, or classic Win32.
        /// </summary>
        static void DetectAndLogAppType(Window window)
        {
            try
            {
                string className = window.Properties.ClassName.IsSupported
                    ? (window.Properties.ClassName.Value ?? "")
                    : "";

                string appType;
                if (className.Contains("ApplicationFrameWindow") || className.Contains("Windows.UI.Core"))
                    appType = "UWP (Modern Windows App)";
                else if (className.StartsWith("pbframe", StringComparison.OrdinalIgnoreCase)
                         || className.StartsWith("PBGUI", StringComparison.OrdinalIgnoreCase)
                         || className.StartsWith("PBGUIObject", StringComparison.OrdinalIgnoreCase))
                    appType = "PowerBuilder";
                else if (className.Contains("WindowsForms"))
                    appType = "Windows Forms (.NET)";
                else if (className.Contains("HwndWrapper") || className.Contains("HwndSource"))
                    appType = "WPF (.NET)";
                else if (className.StartsWith("Chrome_", StringComparison.OrdinalIgnoreCase)
                         || className.Contains("Electron"))
                    appType = "Electron / Chrome-based";
                else
                    appType = "Win32 / Classic";

                Log($"Detected app type: {appType}");
                Log($"Window ClassName : '{className}'");
            }
            catch (Exception ex)
            {
                Log($"Could not detect app type: {ex.Message}");
            }
        }

        // ─── Demo mode (keyboard input) ───────────────────────────────────────────────

        /// <summary>
        /// Focuses the main window of the attached application and types 1234 + 345, then presses Enter.
        /// This is the original calculator demo preserved as its own method.
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
        /// Walks the UI Automation element tree of the main window and prints every accessible element
        /// with its ControlType, Name, AutomationId, and ClassName. Use this to verify which parts of
        /// any application (including PowerBuilder) FlaUI can see and interact with.
        /// </summary>
        /// <param name="mainWindow">The ready main window element returned by <see cref="GetOrLaunchApp"/>.</param>
        static void RunPocInspector(Window mainWindow)
        {
            Log("--- Starting POC UI tree inspection ---");
            Log("This mode prints every accessible UI Automation element in the main window.");
            Log("Elements with a Name or AutomationId can be targeted directly by FlaUI.");
            Log("Elements with no Name and ControlType='Pane' may be opaque (e.g. PowerBuilder DataWindow).");
            Log($"Inspecting window — Title: '{mainWindow.Title}', Class: '{mainWindow.Properties.ClassName.Value}'");

            var sb = new StringBuilder();
            sb.AppendLine("UI Automation Tree");
            sb.AppendLine("==================");

            int nodeCount = PrintUiaTree(mainWindow, sb, indent: 0, maxDepth: 6);
            Console.WriteLine();
            Console.WriteLine(sb.ToString());

            Log($"--- POC inspection completed — {nodeCount} element(s) found ---");
            Log("Review the tree above. Nodes showing Name/AutomationId are directly automatable.");
            Console.WriteLine();
            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }

        static int PrintUiaTree(AutomationElement element, StringBuilder sb, int indent, int maxDepth)
        {
            if (indent > maxDepth)
            {
                sb.AppendLine(Indent(indent) + "... (max depth reached, increase maxDepth if needed)");
                return 0;
            }

            try
            {
                string name = element.Properties.Name.IsSupported ? (element.Properties.Name.Value ?? "") : "(unsupported)";
                string automationId = element.Properties.AutomationId.IsSupported ? (element.Properties.AutomationId.Value ?? "") : "(unsupported)";
                string controlType = element.Properties.ControlType.IsSupported ? element.Properties.ControlType.Value.ToString() : "(unsupported)";
                string className = element.Properties.ClassName.IsSupported ? (element.Properties.ClassName.Value ?? "") : "(unsupported)";

                sb.AppendLine($"{Indent(indent)}[{controlType}] Name=\"{name}\" AutomationId=\"{automationId}\" ClassName=\"{className}\"");

                int count = 1;
                foreach (var child in element.FindAllChildren())
                    count += PrintUiaTree(child, sb, indent + 1, maxDepth);
                return count;
            }
            catch (Exception ex)
            {
                sb.AppendLine($"{Indent(indent)}[ERROR reading element: {ex.Message}]");
                LogError($"Error reading UIA element at depth {indent}", ex);
                return 0;
            }
        }
    }
}
