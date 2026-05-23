using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsCsvExport
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate data with intentional empty cells
            // Row 0
            cells[0, 0].PutValue("ID");
            cells[0, 1].PutValue("Name");
            cells[0, 2].PutValue("Score");
            // Row 1 (empty cell in column B)
            cells[1, 0].PutValue(1);
            // cells[1,1] left empty
            cells[1, 2].PutValue(85);
            // Row 2 (empty row)
            // Row 3 (empty cell in column C)
            cells[3, 0].PutValue(2);
            cells[3, 1].PutValue("Alice");
            // cells[3,2] left empty

            // Configure CSV (text) save options
            TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv)
            {
                // Keep separators for completely blank rows so that column count stays consistent
                KeepSeparatorsForBlankRow = true,
                // Do not trim leading blank rows/columns; keep them as empty strings
                TrimLeadingBlankRowAndColumn = false,
                // Do not trim trailing blank cells in a row
                TrimTailingBlankCells = false,
                // Use ASCII encoding for demonstration (any encoding works)
                Encoding = Encoding.UTF8
            };

            // Save to a memory stream to demonstrate the output without writing a file
            using (MemoryStream ms = new MemoryStream())
            {
                workbook.Save(ms, csvOptions);
                string csvContent = Encoding.UTF8.GetString(ms.ToArray());

                Console.WriteLine("CSV output with empty cells preserved:");
                Console.WriteLine(csvContent);
            }

            // Optionally, save directly to a file
            // workbook.Save("output_preserve_empty.csv", csvOptions);
        }
    }
}