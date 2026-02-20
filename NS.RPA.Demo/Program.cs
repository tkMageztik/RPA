using System;
using System.Diagnostics;
using System.Threading;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Input;
using FlaUI.UIA3;

namespace NS.RPA.Demo
{
    class Program
    {
        private const int CalculatorLoadDelayMs = 2000;

        static void Main(string[] args)
        {
            // Launch calc.exe (on Windows 10+ this starts the UWP Calculator)
            Process.Start("calc.exe");

            // Give the calculator time to fully load
            Thread.Sleep(CalculatorLoadDelayMs);

            using var automation = new UIA3Automation();

            // Find the Calculator process (UWP app on Windows 10+ is "CalculatorApp")
            var calcProcesses = Process.GetProcessesByName("CalculatorApp");
            if (calcProcesses.Length == 0)
                calcProcesses = Process.GetProcessesByName("Calculator");

            if (calcProcesses.Length == 0)
                throw new InvalidOperationException("Calculator process not found. Make sure calc.exe is running.");

            var app = FlaUI.Core.Application.Attach(calcProcesses[0]);

            var mainWindow = app.GetMainWindow(automation);

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
