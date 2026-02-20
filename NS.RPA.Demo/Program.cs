using System;
using System.Diagnostics;
using FlaUI.Core.Input;
using FlaUI.Core.Tools;
using FlaUI.UIA3;

namespace NS.RPA.Demo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Launch calc.exe (on Windows 10+ this starts the UWP Calculator)
            Process.Start("calc.exe");

            using var automation = new UIA3Automation();

            // Wait automatically for the Calculator process to appear — no hard-coded delay needed
            var calcProcess = Retry.WhileNull(
                () =>
                {
                    var procs = Process.GetProcessesByName("CalculatorApp");
                    if (procs.Length == 0)
                        procs = Process.GetProcessesByName("Calculator");
                    return procs.Length > 0 ? procs[0] : null;
                },
                TimeSpan.FromSeconds(30),
                throwOnTimeout: true,
                timeoutMessage: "Calculator process did not start within 30 seconds.");

            var app = FlaUI.Core.Application.Attach(calcProcess);

            // Wait for the main window to be ready (FlaUI polls internally until the timeout)
            var mainWindow = app.GetMainWindow(automation, TimeSpan.FromSeconds(30));

            // Focus the window before typing
            mainWindow.Focus();

            // Type 1234
            Keyboard.Type("1234");

            // Type plus
            Keyboard.Type("+");

            // Type 345
            Keyboard.Type("345");

            // Press Enter to calculate
            Keyboard.Press(FlaUI.Core.WindowsAPI.VirtualKeyShort.RETURN);

            Console.WriteLine("Calculator automation completed: 1234 + 345");
            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }
    }
}
