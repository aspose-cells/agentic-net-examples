using System;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaProtectionCheck
{
    class Program
    {
        static void Main(string[] args)
        {
            // Expect the first argument to be the path of the Excel file (xlsm) to inspect.
            if (args.Length == 0)
            {
                Console.WriteLine("Please provide the path to the Excel file as a command‑line argument.");
                return;
            }

            string filePath = args[0];

            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            // Load the workbook. Aspose.Cells automatically creates a VbaProject object
            // when the workbook is a macro‑enabled file (e.g., .xlsm).
            Workbook workbook = new Workbook(filePath);

            // Access the VBA project.
            VbaProject vbaProject = workbook.VbaProject;

            // Determine whether the VBA project is protected.
            bool isProtected = vbaProject.IsProtected;

            // Prepare a simple object for JSON serialization.
            var result = new
            {
                File = Path.GetFileName(filePath),
                IsVbaProjectProtected = isProtected
            };

            // Serialize the result to JSON.
            string json = JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });

            // Output the JSON string to the console.
            Console.WriteLine(json);
        }
    }
}