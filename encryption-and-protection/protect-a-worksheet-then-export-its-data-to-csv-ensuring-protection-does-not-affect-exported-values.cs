using System;
using Aspose.Cells;

namespace AsposeCellsWorksheetProtectionCsvExport
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (create rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Score");
            sheet.Cells["A2"].PutValue("Alice");
            sheet.Cells["B2"].PutValue(85);
            sheet.Cells["A3"].PutValue("Bob");
            sheet.Cells["B3"].PutValue(92);

            // Protect the worksheet with a password (protect rule)
            sheet.Protect(ProtectionType.All, "myPassword", null);

            // Verify protection (optional)
            Console.WriteLine("Worksheet protected: " + sheet.IsProtected);

            // Export the worksheet data to CSV.
            // The protection does not affect the exported values.
            // Save as CSV (save rule)
            workbook.Save("ProtectedSheetExport.csv", SaveFormat.Csv);

            // Load the CSV back to verify content (load rule)
            Workbook csvLoaded = new Workbook("ProtectedSheetExport.csv");

            // Output the loaded CSV data to console
            Worksheet csvSheet = csvLoaded.Worksheets[0];
            Console.WriteLine("CSV Export Content:");
            for (int row = 0; row <= 3; row++)
            {
                for (int col = 0; col <= 1; col++)
                {
                    Console.Write(csvSheet.Cells[row, col].StringValue + "\t");
                }
                Console.WriteLine();
            }

            // Clean up
            workbook.Dispose();
            csvLoaded.Dispose();
        }
    }
}