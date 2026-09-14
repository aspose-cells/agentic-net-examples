// Title: Copy only visible rows from an Excel worksheet to a new workbook using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens a source .xlsx file, iterates through its rows, skips hidden rows with Cells.IsRowHidden, and copies each visible row to a destination workbook using Cells.CopyRows. | Show how to filter out hidden rows before calling Cells.CopyRows so that only visible rows are duplicated into a new Excel file with Aspose.Cells.
// Common Searches: aspnet copy visible rows Aspose.Cells skip hidden rows | how to copy only non‑hidden rows from one Excel file to another using Aspose.Cells C# | filter hidden rows before using Cells.CopyRows in Aspose.Cells .NET
// Tags: copy visible rows Aspose.Cells | pre-filter hidden rows Aspose.Cells | IsRowHidden method Aspose.Cells | Aspose.Cells copy rows to new workbook | copy rows without hidden rows

using System;
using Aspose.Cells;

namespace AsposeCellsCopyVisibleRows
{
    // The example loads Source.xlsx, loops through each data row, ignores rows where IsRowHidden returns true, copies the remaining visible rows to a new workbook with Cells.CopyRows, and saves the result as Destination.xlsx.
    class Program
    {
        static void Main()
        {
            // Load the source workbook (replace with actual path)
            Workbook sourceWorkbook = new Workbook("Source.xlsx");
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
            Cells sourceCells = sourceSheet.Cells;

            // Create a new workbook for the destination
            Workbook destinationWorkbook = new Workbook();
            Worksheet destinationSheet = destinationWorkbook.Worksheets[0];
            Cells destinationCells = destinationSheet.Cells;

            // Destination row index starts at 0
            int destRowIndex = 0;

            // Iterate through all rows that contain data in the source sheet
            for (int srcRowIndex = 0; srcRowIndex <= sourceCells.MaxDataRow; srcRowIndex++)
            {
                // Skip hidden rows
                if (sourceCells.IsRowHidden(srcRowIndex))
                    continue;

                // Copy a single visible row from source to destination
                // Using CopyRows with rowNumber = 1 copies the row data and formats
                destinationCells.CopyRows(sourceCells, srcRowIndex, destRowIndex, 1);

                // Move to the next destination row
                destRowIndex++;
            }

            // Save the result (replace with desired output path)
            destinationWorkbook.Save("Destination.xlsx", SaveFormat.Xlsx);
        }
    }
}
