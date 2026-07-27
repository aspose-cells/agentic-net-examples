using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Author: Aspose.Cells .NET example – split comma‑separated values in column B into separate columns
    class SplitColumnBByComma
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Example data in column B (index 1) – replace with your own data or load a workbook
            sheet.Cells["B1"].PutValue("Apple,Orange,Banana");
            sheet.Cells["B2"].PutValue("Red,Green,Blue");
            sheet.Cells["B3"].PutValue("Cat,Dog,Mouse");

            // Configure text load options to use comma as the delimiter
            TxtLoadOptions options = new TxtLoadOptions
            {
                Separator = ',' // comma delimiter
            };

            // Determine the number of rows to process (all rows that have data)
            int totalRows = sheet.Cells.MaxDisplayRange.RowCount;

            // Split the text in column B (column index = 1) starting from the first row (row index = 0)
            sheet.Cells.TextToColumns(0, 1, totalRows, options);

            // Save the result
            workbook.Save("SplitColumnB_Output.xlsx");
        }
    }
}