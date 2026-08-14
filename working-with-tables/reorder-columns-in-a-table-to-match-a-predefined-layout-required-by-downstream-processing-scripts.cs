// Title: Reorder Excel worksheet columns and keep ListObject intact with Aspose.Cells for .NET (C#)
// Description: Loads a workbook, creates a new sheet, copies columns to a custom order using Cells.CopyColumns, replicates any existing ListObject (Excel table) with updated headers, removes the original sheet, renames the reordered sheet, and saves the result.
// Keywords: Aspose.Cells C# reorder columns | copy columns Excel Aspose.Cells | preserve ListObject after column reorder | Excel table column order Aspose | worksheet rename Aspose.Cells | copy rows and columns Aspose.Cells | custom column index array Aspose | Excel automation .NET
// Common Searches: Aspose.Cells reorder worksheet columns C# | how to keep Excel table when reordering columns with Aspose | copy columns to new sheet Aspose.Cells .NET | rename sheet after column rearrangement Aspose | preserve ListObject after column copy Aspose.Cells
// Developer Intent: Rearrange specific columns in a worksheet, maintain any embedded Excel table, and replace the original sheet with the reordered version.
// Use Cases: Reorder columns C, A, D, B in a source workbook and export the reordered file. | Move an existing ListObject to a new sheet while changing column order, keeping the table name and headers correct. | Prepare data files for downstream processes that require a fixed column layout, removing the old sheet and renaming the new one.
// AI Prompts: Write C# code with Aspose.Cells that reorders worksheet columns based on an integer array and preserves any ListObject on the sheet. | Show how to copy a table range to a new worksheet, update its column names after reordering, and delete the original worksheet using Aspose.Cells. | Explain how to handle multiple ListObjects on a sheet when reordering columns with Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsColumnReorder
{
    // Loads a workbook, creates a new sheet, copies columns to a custom order using Cells.CopyColumns, replicates any existing ListObject (Excel table) with updated headers, removes the original sheet, renames the reordered sheet, and saves the result.
    class Program
    {
        static void Main()
        {
            // Load the source workbook (replace with your actual file path)
            Workbook workbook = new Workbook("SourceData.xlsx");

            // Assume the data to be reordered is on the first worksheet
            Worksheet sourceSheet = workbook.Worksheets[0];

            // Create a new worksheet that will hold the columns in the required order
            Worksheet reorderedSheet = workbook.Worksheets.Add("Reordered");

            // Define the desired column order (0‑based indexes of the source columns)
            // Example: new order = Column C, Column A, Column D, Column B
            int[] desiredOrder = new int[] { 2, 0, 3, 1 };

            // Copy each column from the source sheet to the new sheet according to the desired order
            for (int destCol = 0; destCol < desiredOrder.Length; destCol++)
            {
                int srcCol = desiredOrder[destCol];
                // Copy a single column (columnNumber = 1) from source to destination
                reorderedSheet.Cells.CopyColumns(
                    sourceSheet.Cells,   // source cells
                    srcCol,              // source column index
                    destCol,             // destination column index
                    1);                  // number of columns to copy
            }

            // If the source sheet contains a ListObject (table), copy it to the new sheet
            // and update its column names so they match the header cells.
            if (sourceSheet.ListObjects.Count > 0)
            {
                // Copy the entire table range (including headers) to the new sheet
                ListObject sourceTable = sourceSheet.ListObjects[0];
                int firstRow = sourceTable.StartRow;
                int firstCol = sourceTable.StartColumn;
                int totalRows = sourceTable.EndRow - firstRow + 1;
                int totalCols = sourceTable.EndColumn - firstCol + 1;

                // Copy the range that contains the table
                reorderedSheet.Cells.CopyRows(
                    sourceSheet.Cells,
                    firstRow,
                    firstRow,
                    totalRows);

                // Re‑create the table on the reordered sheet (same size, with headers)
                int newTableIndex = reorderedSheet.ListObjects.Add(
                    firstRow,
                    0,                     // destination column is now 0 after reordering
                    firstRow + totalRows - 1,
                    desiredOrder.Length - 1,
                    true);
                ListObject newTable = reorderedSheet.ListObjects[newTableIndex];
                newTable.DisplayName = sourceTable.DisplayName;

                // Ensure column names reflect the header cells after reordering
                newTable.UpdateColumnName();
            }

            // Remove the original sheet and rename the reordered sheet to the original name
            int sourceIndex = sourceSheet.Index;
            workbook.Worksheets.RemoveAt(sourceIndex);
            reorderedSheet.Name = "Sheet1";

            // Save the modified workbook (replace with your desired output path)
            workbook.Save("ReorderedData.xlsx");
        }
    }
}
