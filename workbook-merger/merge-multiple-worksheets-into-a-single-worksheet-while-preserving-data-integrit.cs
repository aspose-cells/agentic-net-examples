using System;
using Aspose.Cells;

namespace MergeWorksheetsIntoOne
{
    class Program
    {
        static void Main()
        {
            // Load the source workbook that contains multiple worksheets
            // (replace with your actual file path)
            string sourcePath = "SourceWorkbook.xlsx";
            Workbook sourceWorkbook = new Workbook(sourcePath);

            // Create a new workbook that will hold the merged data
            Workbook destWorkbook = new Workbook();

            // Use the first worksheet of the destination workbook as the target sheet
            Worksheet targetSheet = destWorkbook.Worksheets[0];
            targetSheet.Name = "MergedSheet";

            // Keep track of the next free row in the target sheet
            int targetRowOffset = 0;

            // Iterate through each worksheet in the source workbook
            foreach (Worksheet srcSheet in sourceWorkbook.Worksheets)
            {
                // Determine the used range of the source sheet
                // MaxDataRow/MaxDataColumn give the last row/column that contain data (zero‑based)
                int lastRow = srcSheet.Cells.MaxDataRow;
                int lastColumn = srcSheet.Cells.MaxDataColumn;

                // If the sheet is empty, skip it
                if (lastRow < 0 || lastColumn < 0)
                    continue;

                // Number of rows and columns to copy (add 1 because indexes are zero‑based)
                int rowCount = lastRow + 1;
                int columnCount = lastColumn + 1;

                // Copy rows (data, formatting, formulas) from the source sheet to the target sheet
                // Parameters: source cells, source start row, destination start row, number of rows
                targetSheet.Cells.CopyRows(srcSheet.Cells, 0, targetRowOffset, rowCount);

                // Copy columns (to preserve column widths) from the source sheet to the target sheet
                // Parameters: source cells, source start column, destination start column, number of columns
                targetSheet.Cells.CopyColumns(srcSheet.Cells, 0, 0, columnCount);

                // Optionally, add a blank row between merged sections for readability
                targetRowOffset += rowCount + 1;
            }

            // Save the merged workbook (replace with your desired output path)
            string outputPath = "MergedWorkbook.xlsx";
            destWorkbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Worksheets merged successfully into '{outputPath}'.");
        }
    }
}