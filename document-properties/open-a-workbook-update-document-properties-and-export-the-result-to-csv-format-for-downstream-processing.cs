using System;
using Aspose.Cells;

namespace AsposeCellsDocumentPropertiesToCsv
{
    class Program
    {
        static void Main()
        {
            // Path to the source Excel file
            string sourcePath = "input.xlsx";

            // Path for the resulting CSV file
            string csvPath = "output.csv";

            // Load the workbook from the existing file (load rule)
            Workbook workbook = new Workbook(sourcePath);

            // Update built‑in document properties (property rule)
            // Example: set Author and Title
            workbook.BuiltInDocumentProperties["Author"].Value = "John Doe";
            workbook.BuiltInDocumentProperties["Title"].Value = "Sample Data Export";

            // Add a custom document property (property rule)
            // Example: add a custom property named "ExportedOn"
            workbook.CustomDocumentProperties.Add("ExportedOn", DateTime.Now);

            // Save the workbook as CSV (save rule)
            workbook.Save(csvPath, SaveFormat.Csv);

            // Clean up
            workbook.Dispose();

            Console.WriteLine($"Workbook properties updated and saved as CSV to '{csvPath}'.");
        }
    }
}