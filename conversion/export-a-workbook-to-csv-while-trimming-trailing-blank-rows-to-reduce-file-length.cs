using System;
using System.Text;
using Aspose.Cells;

namespace CsvExportWithTrim
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet and its cells
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate some data
            cells["A1"].PutValue("ID");
            cells["B1"].PutValue("Name");
            cells["A2"].PutValue(1);
            cells["B2"].PutValue("Alice");
            cells["A3"].PutValue(2);
            cells["B3"].PutValue("Bob");

            // Add a few blank rows at the bottom (these will be trimmed)
            // Row 4 and 5 are intentionally left empty
            // Row 6 contains data to demonstrate that only trailing blanks are removed
            cells["A6"].PutValue(3);
            cells["B6"].PutValue("Charlie");

            // Delete all blank rows (trailing blanks will be removed)
            // This uses the provided DeleteBlankRows method.
            cells.DeleteBlankRows();

            // Configure CSV save options
            TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.Csv)
            {
                // Trim trailing blank cells in each row (optional but often desired)
                TrimTailingBlankCells = true,
                // Use UTF-8 encoding for the CSV file
                Encoding = Encoding.UTF8,
                // Set the separator to a comma (standard CSV)
                Separator = ','
            };

            // Save the workbook to a CSV file (lifecycle rule: save)
            workbook.Save("TrimmedOutput.csv", saveOptions);
        }
    }
}