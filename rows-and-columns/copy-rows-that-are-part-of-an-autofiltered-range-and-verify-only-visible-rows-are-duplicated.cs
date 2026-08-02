// Title: Copy Visible Rows from an AutoFiltered Range with Aspose.Cells for .NET (C#)
// Description: Shows how to apply an AutoFilter, detect non‑hidden rows via IsRowHidden, copy only those rows to another worksheet using Cells.CopyRow, and save the result as an XLSX file.
// Keywords: Aspose.Cells | C# | .NET | copy visible rows | AutoFilter | IsRowHidden | Cells.CopyRow | filter rows Excel | duplicate filtered rows | Excel automation | workbook worksheet copy
// Common Searches: Aspose.Cells copy only visible rows | C# copy filtered rows to another sheet | How to duplicate rows after AutoFilter using Aspose | Copy non‑hidden rows Aspose.Cells .NET | Extract filtered rows with Aspose.Cells
// Developer Intent: Transfer only the rows that remain visible after an AutoFilter is applied from a source worksheet to a destination worksheet.
// Use Cases: Generate a report that includes only records meeting specific filter criteria. | Create a clean data export without hidden rows for downstream processing. | Build a summary sheet that mirrors filtered data while preserving original formatting.
// AI Prompts: Provide a C# Aspose.Cells snippet that copies only visible rows from a filtered range to a new worksheet, keeping cell styles intact. | Write a reusable method that accepts source and destination worksheet names and copies rows not hidden by an AutoFilter. | Explain how IsRowHidden and Cells.CopyRow can be combined to duplicate filtered rows in an Excel workbook using Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsCopyVisibleRowsDemo
{
    // Shows how to apply an AutoFilter, detect non‑hidden rows via IsRowHidden, copy only those rows to another worksheet using Cells.CopyRow, and save the result as an XLSX file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet (source)
            Workbook workbook = new Workbook();
            Worksheet srcSheet = workbook.Worksheets[0];
            Cells srcCells = srcSheet.Cells;

            // Populate sample data: header in A1 and values 1..10 in A2:A11
            srcCells["A1"].PutValue("Number");
            for (int i = 0; i < 10; i++)
            {
                srcCells[i + 1, 0].PutValue(i + 1); // Row index i+1, column 0 (A)
            }

            // Apply an AutoFilter to the header row covering column A
            srcSheet.AutoFilter.Range = "A1:A11";

            // Filter to show only numbers greater than 4 and less than 8 (i.e., 5,6,7)
            srcSheet.AutoFilter.Custom(0, FilterOperatorType.GreaterThan, 4);
            srcSheet.AutoFilter.Custom(0, FilterOperatorType.LessThan, 8);
            srcSheet.AutoFilter.Refresh(); // Hide rows that do not meet the criteria

            // Add a destination worksheet where visible rows will be copied
            Worksheet destSheet = workbook.Worksheets.Add("VisibleRowsCopy");
            Cells destCells = destSheet.Cells;

            // Iterate through the data rows and copy only those that are not hidden
            int destRowIndex = 0; // Destination start row
            int firstDataRow = 1; // Data starts at row index 1 (A2)
            int lastDataRow = srcCells.MaxDataRow; // Last row with data

            for (int srcRow = firstDataRow; srcRow <= lastDataRow; srcRow++)
            {
                // Check if the current source row is hidden by the AutoFilter
                if (!srcCells.IsRowHidden(srcRow))
                {
                    // Copy the entire row from source to destination
                    destCells.CopyRow(srcCells, srcRow, destRowIndex);
                    destRowIndex++;
                }
            }

            // OPTIONAL: Verify the copy by printing values from the destination sheet
            Console.WriteLine("Copied visible rows:");
            for (int r = 0; r < destRowIndex; r++)
            {
                Console.WriteLine(destCells[r, 0].StringValue);
            }

            // Save the workbook to a file
            workbook.Save("CopyVisibleRowsResult.xlsx");
        }
    }
}
