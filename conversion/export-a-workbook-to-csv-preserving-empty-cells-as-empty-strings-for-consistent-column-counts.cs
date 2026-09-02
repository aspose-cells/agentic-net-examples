// Title: Export a C# Aspose.Cells workbook to CSV while preserving empty cells as empty strings for consistent column alignment
// AI Prompts: Generate C# code that uses Aspose.Cells to save a workbook as a CSV file, writing blank cells as empty strings so every row retains the same number of columns. | Configure TxtSaveOptions in Aspose.Cells to disable trimming of leading, trailing, and blank rows/columns during CSV export. | Demonstrate exporting a workbook to a CSV string via MemoryStream while keeping separators for completely blank rows.
// Common Searches: Aspose.Cells C# export to CSV keep empty cells | prevent Aspose.Cells from removing blank columns when saving as CSV | TxtSaveOptions KeepSeparatorsForBlankRow example C# | CSV output with consistent column count using Aspose.Cells | how to preserve trailing blank cells in Aspose.Cells CSV export
// Tags: Aspose.Cells CSV export preserve empty cells | TxtSaveOptions KeepSeparatorsForBlankRow | TrimLeadingBlankRowAndColumn false Aspose.Cells | TrimTailingBlankCells false CSV Aspose.Cells | MemoryStream CSV generation C# Aspose.Cells

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsCsvExport
{
    // Creates a workbook, adds data with intentional blanks, sets TxtSaveOptions (KeepSeparatorsForBlankRow = true, TrimLeadingBlankRowAndColumn = false, TrimTailingBlankCells = false, Encoding = ASCII) to retain empty cells, saves to a MemoryStream as CSV, and prints the CSV string where blank cells appear as empty strings, ensuring each row has the same column count.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate some data with intentional empty cells
            cells["A1"].PutValue("Item");
            cells["B1"].PutValue("Price");
            cells["A2"].PutValue("Apple");
            // B2 is left empty on purpose
            cells["B3"].PutValue(2.99);

            // Create CSV save options
            TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv)
            {
                // Keep separators for completely blank rows so that column count stays consistent
                KeepSeparatorsForBlankRow = true,
                // Do not trim leading blank rows/columns; keep them as empty strings
                TrimLeadingBlankRowAndColumn = false,
                // Ensure trailing blank cells are not removed
                TrimTailingBlankCells = false,
                // Use ASCII encoding for simplicity (can be changed as needed)
                Encoding = Encoding.ASCII
            };

            // Save the workbook to a memory stream using the CSV options
            using (MemoryStream ms = new MemoryStream())
            {
                workbook.Save(ms, csvOptions);
                // Convert the stream to a string for display or further processing
                string csvContent = Encoding.ASCII.GetString(ms.ToArray());

                // Output the CSV content
                Console.WriteLine("CSV output with empty cells preserved:");
                Console.WriteLine(csvContent);
            }
        }
    }
}
