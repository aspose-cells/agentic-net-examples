using System;
using Aspose.Cells;

namespace AsposeCellsDocumentPropertiesToCsv
{
    class Program
    {
        static void Main()
        {
            // Path to the source Excel workbook
            string sourcePath = "input.xlsx";

            // Load the workbook from the file system using the constructor that accepts a file name
            Workbook workbook = new Workbook(sourcePath);

            // ----- Update built‑in document properties -----
            // Example: set Author and Title
            workbook.BuiltInDocumentProperties["Author"].Value = "John Doe";
            workbook.BuiltInDocumentProperties["Title"].Value = "Sales Report Q1";

            // ----- Add custom document properties -----
            // Example: add a custom property indicating processing date
            workbook.CustomDocumentProperties.Add("ProcessedDate", DateTime.Now);
            // Example: add a custom numeric property
            workbook.CustomDocumentProperties.Add("RecordCount", 1250);

            // ----- Export the workbook to CSV format -----
            // The Save method with SaveFormat.Csv will generate a CSV file for the first worksheet
            string outputPath = "output.csv";
            workbook.Save(outputPath, SaveFormat.Csv);

            // Clean up
            workbook.Dispose();

            Console.WriteLine($"Workbook '{sourcePath}' processed and saved as CSV to '{outputPath}'.");
        }
    }
}