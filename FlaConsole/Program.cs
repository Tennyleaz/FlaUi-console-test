using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FlaUI.Core.AutomationElements;
using FlaUI.UIA3;

namespace FlaConsole
{
    internal class Program
    {
        private const int MAX_ITEMS = 200;
        private static int itemCount = 0;

        static void Main(string[] args)
        {
            //var app = FlaUI.Core.Application.Launch("notepad.exe");
            var pids = WindowPids.GetPidsWithTopLevelWindows();
            foreach (var p in pids) 
            {

            }

            var app = FlaUI.Core.Application.Attach(35048);
            using (var automation = new UIA3Automation())
            {
                var window = app.GetMainWindow(automation, TimeSpan.FromSeconds(3));
                if (window != null)
                {
                    Console.WriteLine("Title: " + window.Title);
                    PrintUiTree(window, 0, 0, 3);
                }
                else
                {
                    Console.WriteLine("Failed to get the main window.");
                }
            }

            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }

        /// <summary>
        /// DFS to iterate the items of given element.
        /// </summary>
        /// <param name="element"></param>
        /// <param name="index"></param>
        /// <param name="depth"></param>
        /// <param name="maxDepth"></param>
        /// <returns>False if max item/depth reached, true otherwize.</returns>
        private static bool PrintUiTree(AutomationElement element, int index, int depth, int maxDepth)
        {
            if (depth >= maxDepth)
                return false;

            // Print self
            string indent = "";
            for (int i = 0; i < depth; i++)
            {
                indent += "  ";
            }
            // Get automation id
            string id = null;
            try
            {
                id = element.AutomationId;
            }
            catch {}
            // Get name
            string name = null;
            try
            {
                name = element.Name;
            }
            catch { }

            string line = $"{indent}[{element.ControlType} {index}]";
            if (!string.IsNullOrEmpty(name))
                line += $" - \"{element.Name}\"";
            if (!string.IsNullOrEmpty(id))
                line += $" (id={id})";
            if (!element.IsEnabled)
                line += " (Disabled)";

            Console.WriteLine(line);

            itemCount++;
            if (itemCount > MAX_ITEMS)
            {
                Console.WriteLine($"Max {MAX_ITEMS} elements reached, skipping...");
                return false;
            }

            // Print descendants
            var descendants = element.FindAllDescendants();
            for (int i = 0; i < descendants.Length; i++)
            {
                bool isContinue = PrintUiTree(descendants[i], i, depth + 1, maxDepth);
                if (!isContinue)
                    break;
            }

            return true;
        }

        /// <summary>
        /// DFS to find given element by index array.
        /// </summary>
        /// <param name="element"></param>
        /// <param name="index"></param>
        /// <param name="depth"></param>
        /// <param name="indexes"></param>
        /// <returns></returns>
        private static AutomationElement FindElementByIndexSequence(AutomationElement element, int index, int depth, int[] indexes)
        {
            if (depth > indexes.Length)
                return null;

            // Check if element is the right element index
            if (depth == indexes.Length - 1)
            {
                if (index == indexes[depth])
                {
                    return element;
                }
                
                return null;
            }

            // Check descendants
            var descendants = element.FindAllDescendants();
            for (int i = 0; i < descendants.Length; i++)
            {
                AutomationElement found = FindElementByIndexSequence(descendants[i], i, depth + 1, indexes);
                if (found != null)
                    return found;
            }

            return null;
        }
    }
}
