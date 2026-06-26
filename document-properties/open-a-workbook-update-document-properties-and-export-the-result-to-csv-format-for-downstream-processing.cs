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

            // Load the workbook from the file (uses Workbook(string) constructor)
            using (Workbook workbook = new Workbook(sourcePath))
            {
                // Update built‑in document properties
                // Example: set Author and Title
                workbook.BuiltInDocumentProperties["Author"].Value = "John Doe";
                workbook.BuiltInDocumentProperties["Title"].Value = "Processed Data";

                // Add a custom document property
                // If the property already exists, you may update its Value instead
                workbook.CustomDocumentProperties.Add("ProcessedDate", DateTime.Now);

                // Export the workbook to CSV format (uses Save(string, SaveFormat) method)
                // The CSV will contain the data from the first worksheet by default
                string csvOutputPath = "output.csv";
                workbook.Save(csvOutputPath, SaveFormat.Csv);
            }

            Console.WriteLine("Workbook properties updated and saved as CSV successfully.");
        }
    }
}