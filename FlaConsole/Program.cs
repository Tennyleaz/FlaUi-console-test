using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FlaUI.Core.AutomationElements;
using FlaUI.UIA3;
using System.Diagnostics;

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
                    // MaxDepth <= 0 means no explicit depth limit.
                    // In GitHub Desktop, menu bar can be deeper than 5 levels in child-mode traversal.
                    PrintUiTree(window, new List<int> { 0 }, 0, maxDepth: 0);

                    var path = new[] { 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 2 };
                    //var fallback = new List<ElementPathStep>
                    //{
                    //    new ElementPathStep("Window", "GitHub Desktop"),
                    //    new ElementPathStep("Pane", "GitHub Desktop"),
                    //    new ElementPathStep("MenuItem", "View")
                    //};

                    var target = FindElementByIndexSequence(window, path, null, out var foundReason);
                    if (target != null)
                    {
                        Console.WriteLine($"Found element: {DescribeElement(target)}");
                        target.Click();
                    }
                    else
                    {
                        Console.WriteLine($"Not found. Reason: {foundReason}");
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

        /// <summary>
        /// DFS to iterate the items of given element.
        /// </summary>
        /// <param name="element"></param>
        /// <param name="indexPath"></param>
        /// <param name="depth"></param>
        /// <param name="maxDepth"></param>
        /// <returns>False if max item/depth reached, true otherwize.</returns>
        private static bool PrintUiTree(AutomationElement element, List<int> indexPath, int depth, int maxDepth)
        {
            if (maxDepth > 0 && depth >= maxDepth)
            {
                Console.WriteLine($"  {new string(' ', depth * 2)}[...depth limit reached at depth={depth}, path={PrintUiTreeNodePath(indexPath)}...]");
                return false;
            }

            // Print self
            string indent = "";
            for (int i = 0; i < depth; i++)
            {
                indent += "  ";
            }
            string id = SafeGetValue(() => element.AutomationId);
            string name = SafeGetValue(() => element.Name);
            string controlType = GetControlTypeName(element);

            string line = $"{indent}[{controlType} {indexPath.Last()}]";
            if (!string.IsNullOrEmpty(name))
                line += $" - \"{name}\"";
            if (!string.IsNullOrEmpty(id))
                line += $" (id={id})";
            if (!SafeGetValue(() => element.IsEnabled, true))
                line += " (Disabled)";
            line += $" (path={PrintUiTreeNodePath(indexPath)})";

            Console.WriteLine(line);

            itemCount++;
            if (itemCount > MAX_ITEMS)
            {
                Console.WriteLine($"Max {MAX_ITEMS} elements reached, skipping...");
                return false;
            }

            // Print direct children
            var children = SafeGetChildren(element);
            if (children == null)
                return false;

            for (int i = 0; i < children.Length; i++)
            {
                indexPath.Add(i);
                bool isContinue = PrintUiTree(children[i], indexPath, depth + 1, maxDepth);
                indexPath.RemoveAt(indexPath.Count - 1);
                if (!isContinue)
                    break;
            }

            return true;
        }

        /// <summary>
        /// Build index path string for UI tree nodes.
        /// </summary>
        /// <param name="element"></param>
        /// <param name="indexes"></param>
        /// <returns></returns>
        private static string PrintUiTreeNodePath(IReadOnlyList<int> indexPath)
            => indexPath != null ? string.Join(",", indexPath) : string.Empty;

        /// <summary>
        /// Find element by index sequence based on direct children and optional fallback metadata.
        /// </summary>
        /// <param name="rootElement"></param>
        /// <param name="indexes"></param>
        /// <param name="fallbacks"></param>
        /// <param name="reason"></param>
        /// <returns></returns>
        private static AutomationElement FindElementByIndexSequence(AutomationElement rootElement, int[] indexes, IReadOnlyList<ElementPathStep> fallbacks, out string reason)
        {
            reason = null;
            if (rootElement == null)
            {
                reason = "Root element is null.";
                return null;
            }

            if (indexes == null || indexes.Length == 0)
            {
                reason = "Index sequence is empty.";
                return null;
            }

            if (indexes.Length > 0 && indexes[0] != 0)
            {
                var rootFallback = fallbacks != null && fallbacks.Count > 0 ? fallbacks[0] : default(ElementPathStep);
                if (!MatchesElement(rootElement, rootFallback, 0, out var rootMismatch) )
                {
                    reason = $"Depth 0: expected root index 0, got {indexes[0]}. Root mismatch: {rootMismatch}";
                    return null;
                }
            }

            AutomationElement current = rootElement;
            var currentPath = new List<int> { 0 };

            for (int depth = 1; depth < indexes.Length; depth++)
            {
                var children = SafeGetChildren(current, depth, out var childError);
                if (children == null)
                {
                    reason = $"Depth {depth}: unable to get children ({childError}). Path={PrintUiTreeNodePath(currentPath)}";
                    return null;
                }

                int targetIndex = indexes[depth];
                if (targetIndex >= 0 && targetIndex < children.Length)
                {
                    current = children[targetIndex];
                    currentPath.Add(targetIndex);
                    continue;
                }

                var fallback = fallbacks != null && fallbacks.Count > depth ? fallbacks[depth] : default(ElementPathStep);
                var fallbackTarget = FindMatchingChild(children, fallback, depth, out var fallbackReason);
                if (fallbackTarget != null)
                {
                    current = fallbackTarget;
                    currentPath.Add(GetElementIndex(children, current));
                    continue;
                }

                reason = $"Depth {depth}: index {targetIndex} out of range ({children.Length}), and fallback not matched ({fallbackReason}). Path={PrintUiTreeNodePath(currentPath)}";
                return null;
            }

            return current;
        }

        private static int GetElementIndex(AutomationElement[] elements, AutomationElement target)
        {
            for (int i = 0; i < elements.Length; i++)
            {
                if (ReferenceEquals(elements[i], target))
                    return i;
            }
            return -1;
        }

        private static AutomationElement FindMatchingChild(
            AutomationElement[] children,
            ElementPathStep fallback,
            int depth,
            out string reason)
        {
            reason = "No fallback metadata.";
            if (fallback == null)
                return null;

            string mismatchReason = "No fallback metadata.";
            for (int i = 0; i < children.Length; i++)
            {
                if (MatchesElement(children[i], fallback, depth, out var childMismatchReason))
                {
                    reason = null;
                    return children[i];
                }
                mismatchReason = childMismatchReason;
            }

            reason = $"Fallback mismatch at depth {depth}: {mismatchReason}";
            return null;
        }

        private static bool MatchesElement(AutomationElement element, ElementPathStep fallback, int depth, out string mismatchReason)
        {
            mismatchReason = null;
            if (fallback == null)
                return false;

            string controlType = GetControlTypeName(element);
            string name = SafeGetValue(() => element.Name);
            string automationId = SafeGetValue(() => element.AutomationId);

            if (!string.IsNullOrWhiteSpace(fallback.ControlType) &&
                !string.Equals(controlType, fallback.ControlType, StringComparison.OrdinalIgnoreCase))
            {
                mismatchReason = $"controlType expected '{fallback.ControlType}', actual '{controlType}'";
                return false;
            }

            if (!string.IsNullOrWhiteSpace(fallback.AutomationId) &&
                !string.Equals(automationId, fallback.AutomationId, StringComparison.Ordinal))
            {
                mismatchReason = $"automationId expected '{fallback.AutomationId}', actual '{automationId}'";
                return false;
            }

            if (!string.IsNullOrWhiteSpace(fallback.Name) &&
                !string.Equals(name, fallback.Name, StringComparison.Ordinal))
            {
                mismatchReason = $"name expected '{fallback.Name}', actual '{name}'";
                return false;
            }

            if (fallback.IsEnabled.HasValue && SafeGetValue(() => element.IsEnabled) != fallback.IsEnabled.Value)
            {
                mismatchReason = $"isEnabled expected '{fallback.IsEnabled.Value}', actual '{SafeGetValue(() => element.IsEnabled)}'";
                return false;
            }

            return true;
        }

        private static AutomationElement[] SafeGetChildren(AutomationElement element, int depth, out string error)
        {
            error = null;
            try
            {
                return element?.FindAllChildren().ToArray();
            }
            catch (Exception ex)
            {
                error = $"Automation error at depth {depth}: {ex.GetType().Name} - {ex.Message}";
                return null;
            }
        }

        private static string SafeGetValue(Func<string> getter, string defaultValue = null)
        {
            try { return getter(); } catch { return defaultValue; }
        }

        private static bool SafeGetValue(Func<bool> getter, bool defaultValue = true)
        {
            try { return getter(); } catch { return defaultValue; }
        }

        private static string DescribeElement(AutomationElement element)
        {
            return string.IsNullOrWhiteSpace(element?.Name) ? "<no name>" : element.Name;
        }

        private static string GetControlTypeName(AutomationElement element)
            => SafeGetValue(() => element.ControlType.ToString()) ?? "Unknown";

        private static AutomationElement[] SafeGetChildren(AutomationElement element)
            => SafeGetChildren(element, 0, out var _);

        private sealed class ElementPathStep
        {
            public string ControlType { get; }
            public string AutomationId { get; }
            public string Name { get; }
            public int Index { get; }
            public bool? IsEnabled { get; }

            public ElementPathStep(string controlType, string name = null, string automationId = null, int index = 0, bool? isEnabled = null)
            {
                ControlType = controlType;
                Name = name;
                AutomationId = automationId;
                Index = index;
                IsEnabled = isEnabled;
            }
        }
    }
}
