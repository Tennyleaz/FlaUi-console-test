using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FlaUI.UIA3;

namespace FlaConsole
{
    internal class Program
    {
        static void Main(string[] args)
        {
            //var app = FlaUI.Core.Application.Launch("notepad.exe");
            var pids = WindowPids.GetPidsWithTopLevelWindows();
            foreach (var p in pids) 
            { 
            }

            var app = FlaUI.Core.Application.Attach((int)pids.Last());
            using (var automation = new UIA3Automation())
            {
                var window = app.GetMainWindow(automation, TimeSpan.FromSeconds(3));
                if (window != null)
                {
                    Console.WriteLine("Title: " + window.Title);
                    var descendants = window.FindAllDescendants();
                    foreach (var descendant in descendants)
                    {
                        try
                        {
                            Console.WriteLine($"[{descendant.ControlType}] - {descendant.Name} ({descendant.HelpText})");
                        }
                        catch (Exception ex)
                        {
                            //Console.WriteLine($"[{descendant.ControlType}] - {descendant.Name}");
                        }
                    }
                }
                else
                {
                    Console.WriteLine("Failed to get the main window.");
                }
            }

            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }
    }
}
