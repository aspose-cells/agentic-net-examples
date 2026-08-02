using System;
using Aspose.Cells;

namespace AsposeCellsDocumentPropertiesToCsv
{
    class Program
    {
        static void Main()
        {
            // Path to the source Excel file
            string sourceFile = "input.xlsx";

            // Path for the resulting CSV file
            string csvFile = "output.csv";

            // Open the workbook using the constructor that accepts a file path
            using (Workbook workbook = new Workbook(sourceFile))
            {
                // Update a built‑in document property (Author)
                workbook.BuiltInDocumentProperties["Author"].Value = "John Doe";

                // Add a custom document property (ProcessedDate)
                workbook.CustomDocumentProperties.Add("ProcessedDate", DateTime.Now);

                // Save the workbook as CSV using the Save method with SaveFormat
                workbook.Save(csvFile, SaveFormat.Csv);
            }

            Console.WriteLine($"Workbook properties updated and saved as CSV to '{csvFile}'.");
        }
    }
}