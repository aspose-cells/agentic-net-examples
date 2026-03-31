using System;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaProtectionDemo
{
    public class VbaProtectionEvaluator
    {
        /// <summary>
        /// Evaluates the VBA project protection status and locked‑for‑viewing flag of the specified workbook.
        /// Returns a JSON string containing the results.
        /// </summary>
        /// <param name="filePath">Path to the Excel file (xls, xlsx, xlsm, etc.).</param>
        /// <returns>JSON string with IsProtected and IslockedForViewing properties.</returns>
        public static string EvaluateVbaProtection(string filePath)
        {
            // Load the workbook from the given file path
            Workbook workbook = new Workbook(filePath);

            // Access the VBA project associated with the workbook
            VbaProject vbaProject = workbook.VbaProject;

            // Prepare the result object
            var result = new
            {
                IsProtected = vbaProject?.IsProtected ?? false,
                IslockedForViewing = vbaProject?.IslockedForViewing ?? false
            };

            // Serialize the result to JSON
            string jsonResult = JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });

            // Return the JSON string
            return jsonResult;
        }

        // Example usage
        public static void Main(string[] args)
        {
            // Ensure a file path is provided
            if (args.Length == 0)
            {
                Console.WriteLine("Please provide the path to an Excel file as an argument.");
                return;
            }

            string filePath = args[0];

            // Verify the file exists
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            // Evaluate VBA protection and output JSON
            string json = EvaluateVbaProtection(filePath);
            Console.WriteLine(json);
        }
    }
}