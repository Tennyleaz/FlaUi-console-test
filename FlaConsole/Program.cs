using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FlaUI.Core.AutomationElements;
using FlaUI.UIA3;

namespace FlaConsole
{
    internal class Program
    {
        private const int DEFAULT_MAX_ITEMS = 200;
        private static int itemCount = 0;
        private static int printItemLimit = DEFAULT_MAX_ITEMS;

        private enum CliCommand
        {
            List,
            Tree,
            Find,
            Click,
            Subtree,
            Help
        }

        private sealed class CliRequest
        {
            public CliCommand Command;
            public int? Pid;
            public int[] Path;
            public int Depth = 0;
            public int? MaxItems;
            public bool DryRun;
            public bool UseJson;
            public bool ShowHelp;
            public bool RequireMatch;
            public bool UseFallback;
        }

        private sealed class WindowSession : IDisposable
        {
            public FlaUI.Core.Application App { get; }
            public UIA3Automation Automation { get; }
            public AutomationElement Window { get; }

            public WindowSession(FlaUI.Core.Application app, UIA3Automation automation, AutomationElement window)
            {
                App = app;
                Automation = automation;
                Window = window;
            }

            public void Dispose()
            {
                Automation?.Dispose();
                App?.Dispose();
            }
        }

        static int Main(string[] args)
        {
            if (!TryParseArguments(args, out var request, out var parseError))
            {
                Console.WriteLine($"Error: {parseError}");
                PrintGeneralUsage();
                return 1;
            }

            if (request.ShowHelp)
            {
                PrintUsage(request.Command);
                return 0;
            }

            switch (request.Command)
            {
                case CliCommand.List:
                    return HandleList(request);
                case CliCommand.Tree:
                    return HandleTree(request);
                case CliCommand.Find:
                    return HandleFind(request);
                case CliCommand.Click:
                    return HandleClick(request);
                case CliCommand.Subtree:
                    return HandleSubtree(request);
                default:
                    return HandleUnknownOrHelp(request);
            }
        }

        private static int HandleUnknownOrHelp(CliRequest request)
        {
            PrintGeneralUsage();
            return 1;
        }

        private static int HandleList(CliRequest request)
        {
            var windows = WindowPids.GetPidsWithTopLevelWindows();
            if (request.UseJson)
            {
                Console.WriteLine(BuildWindowListJson(windows));
                return 0;
            }

            if (windows.Count == 0)
            {
                Console.WriteLine("No top-level windows found.");
                return 0;
            }

            Console.WriteLine("[pid]\t[handle]\t[process]\t[title]");
            foreach (var item in windows)
            {
                string title = EscapeDisplay(item.Title);
                Console.WriteLine($"{item.Pid}\t{item.Handle}\t{item.ProcessName}\t{title}");
            }
            return 0;
        }

        private static int HandleTree(CliRequest request)
        {
            if (!TryOpenWindow(request.Pid.Value, out var session, out var reason))
            {
                PrintFailure("tree", reason, request.UseJson);
                return 1;
            }

            using (session)
            {
                var oldLimit = printItemLimit;
                if (request.MaxItems.HasValue)
                    printItemLimit = Math.Max(1, request.MaxItems.Value);

                try
                {
                    itemCount = 0;
                    var rootPath = new List<int> { 0 };
                    if (request.UseJson)
                    {
                        var lines = new List<string>();
                        PrintUiTree(session.Window, rootPath, 0, request.Depth, line => lines.Add(line));
                        Console.WriteLine(BuildTreeJson(request.Pid.Value, request.Depth, lines));
                    }
                    else
                    {
                        PrintUiTree(session.Window, rootPath, 0, request.Depth, Console.WriteLine);
                    }
                }
                finally
                {
                    printItemLimit = oldLimit;
                }
            }

            return 0;
        }

        private static int HandleFind(CliRequest request)
        {
            if (!TryOpenWindow(request.Pid.Value, out var session, out var reason))
            {
                PrintFailure("find", reason, request.UseJson, request.Path, request.Pid.Value);
                return 1;
            }

            using (session)
            {
                var element = FindElementByIndexSequence(session.Window, request.Path, out var foundReason);
                if (element == null)
                {
                    PrintFailure("find", foundReason ?? "Element not found.", request.UseJson, request.Path, request.Pid.Value);
                    return 1;
                }

                if (request.UseJson)
                {
                    Console.WriteLine(BuildElementInfoJson(request.Pid.Value, request.Path, element));
                }
                else
                {
                    PrintElementSummary(request.Pid.Value, request.Path, element);
                }
            }
            return 0;
        }

        private static int HandleSubtree(CliRequest request)
        {
            if (!TryOpenWindow(request.Pid.Value, out var session, out var reason))
            {
                PrintFailure("subtree", reason, request.UseJson, request.Path, request.Pid.Value);
                return 1;
            }

            using (session)
            {
                var element = FindElementByIndexSequence(session.Window, request.Path, out var foundReason);
                if (element == null)
                {
                    PrintFailure("subtree", foundReason ?? "Element not found.", request.UseJson, request.Path, request.Pid.Value);
                    return 1;
                }

                var oldLimit = printItemLimit;
                if (request.MaxItems.HasValue)
                    printItemLimit = Math.Max(1, request.MaxItems.Value);

                try
                {
                    itemCount = 0;
                    var rootPath = new List<int>(request.Path);
                    if (request.UseJson)
                    {
                        var lines = new List<string>();
                        PrintUiTree(element, rootPath, 0, request.Depth, lines.Add);
                        Console.WriteLine(BuildSubTreeJson(request.Pid.Value, request.Path, request.Depth, lines));
                    }
                    else
                    {
                        if (!request.UseJson)
                        {
                            var pathText = PrintUiTreeNodePath(request.Path);
                            Console.WriteLine($"[subtree root={pathText}]");
                        }
                        PrintUiTree(element, rootPath, 0, request.Depth, Console.WriteLine);
                    }
                }
                finally
                {
                    printItemLimit = oldLimit;
                }
            }

            return 0;
        }

        private static int HandleClick(CliRequest request)
        {
            if (!TryOpenWindow(request.Pid.Value, out var session, out var reason))
            {
                PrintFailure("click", reason, request.UseJson, request.Path, request.Pid.Value);
                return 1;
            }

            using (session)
            {
                var element = FindElementByIndexSequence(session.Window, request.Path, out var foundReason);
                if (element == null)
                {
                    PrintFailure("click", foundReason ?? "Element not found.", request.UseJson, request.Path, request.Pid.Value);
                    return 1;
                }

                if (request.UseJson)
                {
                    if (request.DryRun)
                    {
                        Console.WriteLine(BuildElementInfoJson(request.Pid.Value, request.Path, element));
                        return 0;
                    }

                    try
                    {
                        element.Click();
                        Console.WriteLine(BuildClickResultJson(request.Pid.Value, request.Path, true, "clicked"));
                    }
                    catch (Exception ex)
                    {
                        PrintFailure("click", $"Click failed: {ex.GetType().Name}: {ex.Message}", request.UseJson, request.Path, request.Pid.Value);
                        return 1;
                    }
                }
                else
                {
                    PrintElementSummary(request.Pid.Value, request.Path, element);
                    if (request.DryRun)
                    {
                        Console.WriteLine("Dry-run enabled, skip click.");
                        return 0;
                    }

                    try
                    {
                        element.Click();
                        Console.WriteLine("Click executed.");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Click failed: {ex.GetType().Name}: {ex.Message}");
                        return 1;
                    }
                }
            }

            return 0;
        }

        private static bool TryParseArguments(string[] args, out CliRequest request, out string error)
        {
            request = new CliRequest();
            error = null;

            if (args == null || args.Length == 0)
            {
                error = "No command specified.";
                return false;
            }

            if (!TryParseCommand(args[0], out var command))
            {
                error = $"Unknown command '{args[0]}'.";
                return false;
            }

            request.Command = command;
            if (command == CliCommand.Help)
            {
                request.ShowHelp = true;
                return true;
            }

            for (int i = 1; i < args.Length; i++)
            {
                var token = args[i];
                if (string.IsNullOrWhiteSpace(token))
                    continue;

                if (!token.StartsWith("--", StringComparison.Ordinal))
                {
                    error = $"Unexpected argument '{token}'.";
                    return false;
                }

                var payload = token.Substring(2);
                var eqIndex = payload.IndexOf('=');
                var option = eqIndex >= 0 ? payload.Substring(0, eqIndex) : payload;
                var value = eqIndex >= 0 ? payload.Substring(eqIndex + 1) : null;

                if (value == null && (
                    option == "help" ||
                    option == "dry-run" ||
                    option == "json" ||
                    option == "require-match" ||
                    option == "fallback"))
                {
                    value = "true";
                }

                if (!string.IsNullOrEmpty(option) && value == null && i + 1 < args.Length && !args[i + 1].StartsWith("--", StringComparison.Ordinal))
                {
                    value = args[i + 1];
                    i++;
                }

                switch (option.ToLowerInvariant())
                {
                    case "pid":
                        if (string.IsNullOrWhiteSpace(value) || !int.TryParse(value, out var pid) || pid <= 0)
                        {
                            error = "Invalid --pid, expect a positive integer.";
                            return false;
                        }
                        request.Pid = pid;
                        break;
                    case "path":
                        if (!TryParsePath(value, out var path, out var pathError))
                        {
                            error = pathError;
                            return false;
                        }
                        request.Path = path;
                        break;
                    case "depth":
                        if (string.IsNullOrWhiteSpace(value) || !int.TryParse(value, out var depth) || depth < 0)
                        {
                            error = "Invalid --depth, expect an integer >= 0. (0 means no depth limit)";
                            return false;
                        }
                        request.Depth = depth;
                        break;
                    case "max-items":
                        if (string.IsNullOrWhiteSpace(value) || !int.TryParse(value, out var maxItems) || maxItems <= 0)
                        {
                            error = "Invalid --max-items, expect an integer > 0.";
                            return false;
                        }
                        request.MaxItems = maxItems;
                        break;
                    case "dry-run":
                        request.DryRun = true;
                        break;
                    case "json":
                        request.UseJson = true;
                        break;
                    case "require-match":
                        request.RequireMatch = true;
                        break;
                    case "fallback":
                        request.UseFallback = true;
                        break;
                    case "help":
                        request.ShowHelp = true;
                        break;
                    default:
                        error = $"Unknown option '--{option}'.";
                        return false;
                }
            }

            if (request.ShowHelp)
            {
                return true;
            }

            if (request.Command == CliCommand.List)
                return true;

            if (!request.Pid.HasValue)
            {
                error = "--pid is required.";
                return false;
            }

            if ((request.Command == CliCommand.Find || request.Command == CliCommand.Click || request.Command == CliCommand.Subtree) && (request.Path == null || request.Path.Length == 0))
            {
                error = "--path is required.";
                return false;
            }

            return true;
        }

        private static bool TryParseCommand(string input, out CliCommand command)
        {
            command = CliCommand.List;
            if (string.IsNullOrWhiteSpace(input))
                return false;

            switch (input.ToLowerInvariant())
            {
                case "list":
                    command = CliCommand.List;
                    return true;
                case "tree":
                    command = CliCommand.Tree;
                    return true;
                case "find":
                    command = CliCommand.Find;
                    return true;
                case "click":
                    command = CliCommand.Click;
                    return true;
                case "subtree":
                    command = CliCommand.Subtree;
                    return true;
                case "help":
                    command = CliCommand.Help;
                    return true;
                default:
                    return false;
            }
        }

        private static string BuildSubTreeJson(int pid, int[] rootPath, int depth, IReadOnlyList<string> lines)
        {
            var sb = new StringBuilder();
            sb.Append("{\"command\":\"subtree\",");
            sb.Append("\"pid\":").Append(pid).Append(",");
            sb.Append("\"rootPath\":\"").Append(EscapeJson(PrintUiTreeNodePath(rootPath))).Append("\",");
            sb.Append("\"depth\":").Append(depth).Append(",");
            sb.Append("\"items\":[");
            for (int i = 0; i < lines.Count; i++)
            {
                if (i > 0)
                    sb.Append(',');
                sb.Append('"').Append(EscapeJson(lines[i])).Append('"');
            }
            sb.Append("]}");
            return sb.ToString();
        }

        private static bool TryParsePath(string rawPath, out int[] path, out string error)
        {
            error = null;
            path = null;

            if (string.IsNullOrWhiteSpace(rawPath))
            {
                error = "--path cannot be empty.";
                return false;
            }

            var pieces = rawPath.Split(',');
            var indexes = new List<int>(pieces.Length);
            foreach (var pieceRaw in pieces)
            {
                var piece = pieceRaw?.Trim();
                if (string.IsNullOrWhiteSpace(piece) || !int.TryParse(piece, out var index) || index < 0)
                {
                    error = $"--path has invalid segment '{piece}'.";
                    return false;
                }
                indexes.Add(index);
            }

            path = indexes.ToArray();
            return path.Length > 0;
        }

        private static bool TryOpenWindow(int pid, out WindowSession session, out string reason)
        {
            session = null;
            reason = null;

            try
            {
                var app = FlaUI.Core.Application.Attach(pid);
                var automation = new UIA3Automation();
                var window = app.GetMainWindow(automation, TimeSpan.FromSeconds(3));
                if (window == null)
                {
                    automation.Dispose();
                    app.Dispose();
                    reason = $"No main window available for pid {pid}.";
                    return false;
                }

                session = new WindowSession(app, automation, window);
                return true;
            }
            catch (Exception ex)
            {
                reason = $"Attach/GetMainWindow failed: {ex.GetType().Name}: {ex.Message}";
                return false;
            }
        }

        private static void PrintFailure(string command, string reason, bool useJson, int[] path = null, int? pid = null)
        {
            if (useJson)
            {
                Console.WriteLine(BuildFailureJson(command, path, pid, reason));
                return;
            }

            Console.WriteLine($"[{command}] {reason}");
        }

        private static ElementSummary BuildElementSummary(int pid, int[] path, AutomationElement element)
        {
            if (element == null)
            {
                return new ElementSummary
                {
                    Pid = pid,
                    Path = PrintUiTreeNodePath(path),
                    ControlType = string.Empty,
                    Name = string.Empty,
                    AutomationId = string.Empty,
                    Enabled = false
                };
            }

            return new ElementSummary
            {
                Pid = pid,
                Path = PrintUiTreeNodePath(path),
                ControlType = GetControlTypeName(element),
                Name = SafeGetValue(() => element.Name),
                AutomationId = SafeGetValue(() => element.AutomationId),
                Enabled = SafeGetValue(() => element.IsEnabled, true)
            };
        }

        private static void PrintElementSummary(int pid, int[] path, AutomationElement element)
        {
            var summary = BuildElementSummary(pid, path, element);
            Console.WriteLine("Found element:");
            Console.WriteLine($"  pid: {summary.Pid}");
            Console.WriteLine($"  path: {summary.Path}");
            Console.WriteLine($"  controlType: {summary.ControlType}");
            Console.WriteLine($"  name: \"{summary.Name}\"");
            Console.WriteLine($"  automationId: {summary.AutomationId}");
            Console.WriteLine($"  enabled: {summary.Enabled}");
        }

        private static string BuildElementInfoJson(int pid, int[] path, AutomationElement element)
        {
            var summary = BuildElementSummary(pid, path, element);
            var sb = new StringBuilder();
            sb.Append("{");
            sb.Append("\"command\":\"find\",");
            sb.Append("\"success\":true,");
            sb.Append(BuildElementSummaryJson(summary));
            sb.Append("}");
            return sb.ToString();
        }

        private static string BuildClickResultJson(int pid, int[] path, bool success, string note)
        {
            var sb = new StringBuilder();
            sb.Append("{");
            sb.Append("\"command\":\"click\",");
            sb.Append("\"success\":").Append(success.ToString().ToLowerInvariant()).Append(',');
            sb.Append("\"note\":\"").Append(EscapeJson(note)).Append("\",");
            sb.Append(BuildElementSummaryJson(BuildElementSummary(pid, path, null)));
            sb.Append("}");
            return sb.ToString();
        }

        private static string BuildFailureJson(string command, int[] path, int? pid, string reason)
        {
            var sb = new StringBuilder();
            sb.Append("{");
            sb.Append("\"command\":\"").Append(EscapeJson(command)).Append("\",");
            sb.Append("\"success\":false,");
            sb.Append("\"reason\":\"").Append(EscapeJson(reason ?? "Unknown error")).Append("\"");
            if (pid.HasValue)
            {
                sb.Append(",\"pid\":").Append(pid.Value);
            }
            if (path != null)
            {
                sb.Append(",\"path\":\"").Append(EscapeJson(PrintUiTreeNodePath(path))).Append("\"");
            }
            sb.Append("}");
            return sb.ToString();
        }

        private static string BuildElementSummaryJson(ElementSummary summary)
        {
            var sb = new StringBuilder();
            sb.Append("\"pid\":").Append(summary.Pid).Append(',');
            sb.Append("\"path\":\"").Append(EscapeJson(summary.Path)).Append("\",");
            sb.Append("\"controlType\":\"").Append(EscapeJson(summary.ControlType)).Append("\",");
            sb.Append("\"name\":\"").Append(EscapeJson(summary.Name)).Append("\",");
            sb.Append("\"automationId\":\"").Append(EscapeJson(summary.AutomationId)).Append("\",");
            sb.Append("\"enabled\":").Append(summary.Enabled.ToString().ToLowerInvariant());
            return sb.ToString();
        }

        private static string BuildWindowListJson(IEnumerable<WindowInfo> windows)
        {
            var sb = new StringBuilder();
            sb.Append("{\"command\":\"list\",\"windows\":[");
            bool firstWindow = true;
            foreach (var window in windows)
            {
                if (!firstWindow)
                    sb.Append(',');
                firstWindow = false;

                sb.Append("{");
                sb.Append("\"pid\":").Append(window.Pid).Append(',');
                sb.Append("\"handle\":\"").Append(EscapeJson(window.Handle.ToString())).Append("\",");
                sb.Append("\"process\":\"").Append(EscapeJson(window.ProcessName)).Append("\",");
                sb.Append("\"title\":\"").Append(EscapeJson(window.Title)).Append("\"");
                sb.Append("}");
            }
            sb.Append("]}");
            return sb.ToString();
        }

        private static string BuildTreeJson(int pid, int depth, IReadOnlyList<string> lines)
        {
            var sb = new StringBuilder();
            sb.Append("{\"command\":\"tree\",\"pid\":").Append(pid).Append(",\"depth\":").Append(depth).Append(",\"items\":[");
            for (int i = 0; i < lines.Count; i++)
            {
                if (i > 0)
                    sb.Append(',');
                sb.Append('"').Append(EscapeJson(lines[i])).Append('"');
            }
            sb.Append("]}");
            return sb.ToString();
        }

        private static void PrintGeneralUsage()
        {
            PrintUsage(CliCommand.Help);
        }

        private static void PrintUsage(CliCommand command)
        {
            switch (command)
            {
                case CliCommand.List:
                    Console.WriteLine("Usage: flaui list [--json]");
                    break;
                case CliCommand.Tree:
                    Console.WriteLine("Usage: flaui tree --pid <pid> [--depth <n>] [--max-items <n>] [--json]");
                    break;
                case CliCommand.Find:
                    Console.WriteLine("Usage: flaui find --pid <pid> --path <i1,i2,...> [--require-match] [--fallback] [--json]");
                    break;
                case CliCommand.Click:
                    Console.WriteLine("Usage: flaui click --pid <pid> --path <i1,i2,...> [--dry-run] [--json]");
                    break;
                case CliCommand.Subtree:
                    Console.WriteLine("Usage: flaui subtree --pid <pid> --path <i1,i2,...> [--depth <n>] [--max-items <n>] [--json]");
                    break;
                default:
                    Console.WriteLine("Usage:");
                    Console.WriteLine("  flaui list [--json]");
                    Console.WriteLine("  flaui tree --pid <pid> [--depth <n>] [--max-items <n>] [--json]");
                    Console.WriteLine("  flaui find --pid <pid> --path <i1,i2,...> [--require-match] [--fallback] [--json]");
                    Console.WriteLine("  flaui subtree --pid <pid> --path <i1,i2,...> [--depth <n>] [--max-items <n>] [--json]");
                    Console.WriteLine("  flaui click --pid <pid> --path <i1,i2,...> [--dry-run] [--json]");
                    Console.WriteLine();
                    Console.WriteLine("Path uses comma-separated indexes and index 0 is the root.");
                    break;
            }
        }

        private static string EscapeDisplay(string value)
        {
            if (value == null)
                return string.Empty;
            return value.Replace('\t', ' ').Replace('\r', ' ').Replace('\n', ' ');
        }

        private static string EscapeJson(string value)
        {
            if (value == null)
                return string.Empty;
            return value.Replace("\\", "\\\\").Replace("\"", "\\\"").Replace("\r", "\\r").Replace("\n", "\\n");
        }

        /// <summary>
        /// DFS to iterate the items of given element.
        /// </summary>
        /// <param name="element"></param>
        /// <param name="indexPath"></param>
        /// <param name="depth"></param>
        /// <param name="maxDepth"></param>
        /// <param name="writeLine">Callback function to print a line</param>
        /// <returns>False if max item/depth reached, true otherwize.</returns>
        private static bool PrintUiTree(AutomationElement element, List<int> indexPath, int depth, int maxDepth, Action<string> writeLine)
        {
            if (maxDepth > 0 && depth >= maxDepth)
            {
                writeLine?.Invoke($"  {new string(' ', depth * 2)}[...depth limit reached at depth={depth}, path={PrintUiTreeNodePath(indexPath)}...]");
                return false;
            }

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

            writeLine?.Invoke(line);

            itemCount++;
            if (itemCount > printItemLimit)
            {
                writeLine?.Invoke($"Max {printItemLimit} elements reached, skipping...");
                return false;
            }

            var children = SafeGetChildren(element);
            if (children == null)
                return false;

            for (int i = 0; i < children.Length; i++)
            {
                indexPath.Add(i);
                bool isContinue = PrintUiTree(children[i], indexPath, depth + 1, maxDepth, writeLine);
                indexPath.RemoveAt(indexPath.Count - 1);
                if (!isContinue)
                    break;
            }

            return true;
        }

        /// <summary>
        /// Build index path string for UI tree nodes.
        /// </summary>
        /// <returns></returns>
        private static string PrintUiTreeNodePath(IReadOnlyList<int> indexPath)
            => indexPath != null ? string.Join(",", indexPath) : string.Empty;

        /// <summary>
        /// Find element by index sequence based on direct children and optional fallback metadata.
        /// </summary>
        private static AutomationElement FindElementByIndexSequence(AutomationElement rootElement, int[] indexes, out string reason)
        {
            return FindElementByIndexSequence(rootElement, indexes, null, out reason);
        }

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

        private sealed class ElementSummary
        {
            public int Pid;
            public string Path;
            public string ControlType;
            public string Name;
            public string AutomationId;
            public bool Enabled;
        }
    }
}
