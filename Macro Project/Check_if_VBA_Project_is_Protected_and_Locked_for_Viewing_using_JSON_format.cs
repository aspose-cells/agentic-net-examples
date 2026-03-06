using System;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaCheck
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel file (macro-enabled workbook)
            string filePath = "sample.xlsm";

            // Load the workbook
            Workbook workbook = new Workbook(filePath);

            // Access the VBA project
            VbaProject vbaProject = workbook.VbaProject;

            // Retrieve protection status
            bool isProtected = vbaProject.IsProtected;
            bool isLockedForViewing = vbaProject.IslockedForViewing;

            // Create an anonymous object for JSON serialization
            var result = new
            {
                IsProtected = isProtected,
                IsLockedForViewing = isLockedForViewing
            };

            // Serialize to JSON
            string json = JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });

            // Output the JSON string
            Console.WriteLine(json);
        }
    }
}