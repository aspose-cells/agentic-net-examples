using System;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaCheck
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the macro-enabled workbook (can be passed as a command‑line argument)
            string workbookPath = args.Length > 0 ? args[0] : "example.xlsm";

            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"File not found: {workbookPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Access the VBA project (may be null if the workbook has no VBA)
            VbaProject vbaProject = workbook.VbaProject;

            bool isSigned = false;
            bool isValidSigned = false;

            if (vbaProject != null)
            {
                isSigned = vbaProject.IsSigned;
                isValidSigned = vbaProject.IsValidSigned;
            }

            // Prepare result object for JSON serialization
            var result = new
            {
                IsSigned = isSigned,
                IsValidSigned = isValidSigned
            };

            // Serialize result to JSON
            string json = JsonSerializer.Serialize(result);

            // Output JSON string
            Console.WriteLine(json);
        }
    }
}