using System;
using System.Diagnostics;
using System.IO;
using FlaUI.Core.Input;
using FlaUI.Core.Tools;
using FlaUI.UIA3;

namespace NS.RPA.Demo
{
    class Program
    {
        static void Main(string[] args)
        {
            // The application path can be a full path (e.g. "C:\Apps\MyPBApp.exe") or just an
            // executable name (e.g. "calc.exe"). Defaults to "calc.exe" when no argument is given.
            string appPath = args.Length > 0 ? args[0] : "calc.exe";

            if (string.IsNullOrWhiteSpace(appPath))
                throw new ArgumentException("Application path must not be empty.", nameof(args));

            // Derive the process name from the executable file name (without extension),
            // which is what Process.GetProcessesByName expects.
            string processName = Path.GetFileNameWithoutExtension(appPath);

            using var automation = new UIA3Automation();

            // Check whether the application is already running.
            var runningProcesses = Process.GetProcessesByName(processName);

            FlaUI.Core.Application app;
            if (runningProcesses.Length > 0)
            {
                // Attach to the already-running instance instead of opening a second one.
                Console.WriteLine($"'{processName}' is already running — attaching to existing instance.");
                app = FlaUI.Core.Application.Attach(runningProcesses[0]);
            }
            else
            {
                // Launch the application from the provided path.
                Console.WriteLine($"Launching '{appPath}'...");
                try
                {
                    Process.Start(appPath);
                }
                catch (Exception ex)
                {
                    throw new InvalidOperationException(
                        $"Failed to launch '{appPath}'. Verify the path is correct and accessible. Details: {ex.Message}", ex);
                }

                // Wait reactively until the process appears — no hard-coded delay needed.
                var launchedProcess = Retry.WhileNull(
                    () =>
                    {
                        var procs = Process.GetProcessesByName(processName);
                        return procs.Length > 0 ? procs[0] : null;
                    },
                    TimeSpan.FromSeconds(30),
                    throwOnTimeout: true,
                    timeoutMessage: $"Process '{processName}' did not start within 30 seconds.");

                app = FlaUI.Core.Application.Attach(launchedProcess);
            }

            // Wait for the main window to be ready (FlaUI polls internally until the timeout).
            var mainWindow = app.GetMainWindow(automation, TimeSpan.FromSeconds(30));

            // Focus the window before typing.
            mainWindow.Focus();

            // Type 1234
            Keyboard.Type("1234");

            // Type plus
            Keyboard.Type("+");

            // Type 345
            Keyboard.Type("345");

            // Press Enter to calculate
            Keyboard.Press(FlaUI.Core.WindowsAPI.VirtualKeyShort.RETURN);

            Console.WriteLine($"Automation completed for '{processName}'.");
            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }
    }
}
