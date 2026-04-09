using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ScannerApp
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // Check for --list flag first
            if (HasFlag(args, "--list"))
            {
                string listOutput = GetArgValue(args, "--output", "");
                ListScanners(listOutput);
                return;
            }

            // Parse command line arguments
            int sourceIndex = GetArgValue(args, "--sourceindex", 0);
            string output = GetArgValue(args, "--output", "");
            bool feeder = GetArgValue(args, "--feeder", false);
            bool duplex = GetArgValue(args, "--duplex", false);
            string colorMode = GetArgValue(args, "--colormode", "color");
            string resolution = GetArgValue(args, "--resolution", "medium");
            int pageWidth = GetArgValue(args, "--pagewidth", 8500);
            int pageHeight = GetArgValue(args, "--pageheight", 11000);

            // If output is empty, generate default path; if no extension, default to .pdf
            string outPdf = string.IsNullOrEmpty(output)
                ? Path.Combine("c:\\temp\\ScannedImages", $"twain_{DateTime.Now:yyyyMMddHHmmssfff}.pdf")
                : string.IsNullOrEmpty(Path.GetExtension(output)) ? output + ".pdf" : output;

            // Initialize logging with the same path as output PDF but with .log extension
            Logger.Initialize(outPdf);

            // Display parsed values
            //Console.WriteLine($"Source Index: {sourceIndex}");
            //Console.WriteLine($"Output: {outPdf}");
            //Console.WriteLine($"Feeder: {feeder}");
            //Console.WriteLine($"Duplex: {duplex}");
            //Console.WriteLine($"Color Mode: {colorMode}");
            //Console.WriteLine($"Resolution: {resolution}");
            //Console.WriteLine($"Page Width: {pageWidth}");
            //Console.WriteLine($"Page Height: {pageHeight}");

            // Log parsed values
            Logger.Log($"Source Index: {sourceIndex}");
            Logger.Log($"Output: {outPdf}");
            Logger.Log($"Feeder: {feeder}");
            Logger.Log($"Duplex: {duplex}");
            Logger.Log($"Color Mode: {colorMode}");
            Logger.Log($"Resolution: {resolution}");
            Logger.Log($"Page Width: {pageWidth}");
            Logger.Log($"Page Height: {pageHeight}");

            // Call ScanToPdf
            var scanner = new TwainScanner();
            string result = scanner.ScanToPdf(sourceIndex, outPdf, feeder, duplex, colorMode, resolution, pageWidth, pageHeight);

            // Output result as JSON to file
            bool success = string.IsNullOrEmpty(result);
            string json = $"{{\"success\": {success.ToString().ToLower()}, \"message\": \"{EscapeJson(result)}\", \"output\": \"{EscapeJson(outPdf)}\", \"scannerName\": \"{EscapeJson(scanner.ScannerName)}\", \"scannerModel\": \"{EscapeJson(scanner.ScannerModel)}\"}}";
            string jsonPath = Path.ChangeExtension(outPdf, ".json");
            File.WriteAllText(jsonPath, json);
            Logger.Log($"Result: {json}");
        }

        static string EscapeJson(string value)
        {
            if (string.IsNullOrEmpty(value)) return "";
            return value.Replace("\\", "\\\\").Replace("\"", "\\\"").Replace("\n", "\\n").Replace("\r", "\\r").Replace("\t", "\\t");
        }

        static string GetArgValue(string[] args, string key, string defaultValue)
        {
            for (int i = 0; i < args.Length - 1; i++)
            {
                if (args[i].Equals(key, StringComparison.OrdinalIgnoreCase))
                {
                    return args[i + 1];
                }
            }
            return defaultValue;
        }

        static int GetArgValue(string[] args, string key, int defaultValue)
        {
            string value = GetArgValue(args, key, "");
            return int.TryParse(value, out int result) ? result : defaultValue;
        }

        static bool GetArgValue(string[] args, string key, bool defaultValue)
        {
            string value = GetArgValue(args, key, "");
            return bool.TryParse(value, out bool result) ? result : defaultValue;
        }

        static bool HasFlag(string[] args, string flag)
        {
            return args.Any(a => a.Equals(flag, StringComparison.OrdinalIgnoreCase));
        }

        static void ListScanners(string output = "")
        {
            // Initialize logging for list operation
            string logPath = string.IsNullOrEmpty(output)
                ? Path.Combine("c:\\temp\\ScannedImages", $"twain_list_{DateTime.Now:yyyyMMddHHmmssfff}.log")
                : output;
            Logger.Initialize(logPath);
            
            var scanner = new TwainScanner();
            string scannerList = scanner.GetAvailableScanners();
            
            // Parse the scanner list and build JSON array
            var scanners = new List<string>();
            var lines = scannerList.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var line in lines)
            {
                var parts = line.Split(new[] { '=' }, 2);
                if (parts.Length == 2)
                {
                    string index = parts[0].Trim();
                    string name = EscapeJson(parts[1].Trim());
                    scanners.Add($"{{\"index\": {index}, \"name\": \"{name}\"}}");
                }
            }
            
            string json = $"{{\"success\": true, \"scanners\": [{string.Join(", ", scanners)}]}}";
            Console.WriteLine(json);
            Logger.Log($"Result: {json}");
        }
    }
}
