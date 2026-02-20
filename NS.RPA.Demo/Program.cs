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
        /// the given path and waits reactively until its process and main window are ready.
        /// Returns both the FlaUI Application handle and the ready main window element so callers
        /// do not need to call GetMainWindow a second time.
        /// </summary>
        /// <param name="appPath">Full path or bare executable name (e.g. "C:\Apps\MyApp.exe" or "calc.exe").</param>
        /// <param name="automation">The UIA3 automation instance used to locate the main window.</param>
        /// <returns>A tuple containing the attached <see cref="Application"/> and its ready <see cref="Window"/> main window.</returns>
        static (Application app, Window mainWindow) GetOrLaunchApp(string appPath, UIA3Automation automation)
        {
            string processName = Path.GetFileNameWithoutExtension(appPath);
            Log($"Looking for running process '{processName}'...");

            var runningProcesses = Process.GetProcessesByName(processName);

            Application app;
            if (runningProcesses.Length > 0)
            {
                Log($"Process '{processName}' is already running (PID {runningProcesses[0].Id}) — attaching to existing instance.");
                app = Application.Attach(runningProcesses[0]);
            }
            else
            {
                Log($"Process '{processName}' not running — launching '{appPath}'...");
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

                Log($"Waiting for process '{processName}' to appear in the process list...");
                var retryResult = Retry.WhileNull(
                    () =>
                    {
                        var procs = Process.GetProcessesByName(processName);
                        return procs.Length > 0 ? procs[0] : null;
                    },
                    TimeSpan.FromSeconds(30),
                    throwOnTimeout: true,
                    timeoutMessage: $"Process '{processName}' did not appear within 30 seconds.");

                // throwOnTimeout:true guarantees Result is non-null here; the ! suppresses the nullable warning.
                var launchedProcess = retryResult.Result!;
                Log($"Process '{processName}' is now running (PID {launchedProcess.Id}).");
                app = Application.Attach(launchedProcess);
            }

            Log("Waiting for the main window to become ready...");
            var mainWindow = app.GetMainWindow(automation, TimeSpan.FromSeconds(30));
            if (mainWindow == null)
                throw new InvalidOperationException($"Main window of '{processName}' was not found within 30 seconds.");
            Log($"Main window ready — Title: '{mainWindow.Title}', Handle: {mainWindow.Properties.NativeWindowHandle.Value}");

            return (app, mainWindow);
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
