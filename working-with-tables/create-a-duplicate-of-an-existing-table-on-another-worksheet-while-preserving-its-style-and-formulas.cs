// Title: How to duplicate an Excel ListObject to another worksheet while keeping formulas and formatting with Aspose.Cells for .NET
// AI Prompts: Generate C# code that copies a ListObject from one worksheet to a new worksheet, preserving cell values, formulas, and styles using Aspose.Cells. | Show how to create a new worksheet, copy the source table range cell‑by‑cell, and add a new ListObject that retains the original table's formatting. | Provide a step‑by‑step example for cloning an Excel table to a different sheet while maintaining all formulas and formatting with Aspose.Cells.
// Common Searches: Aspose.Cells copy Excel table to another sheet with formulas C# | duplicate ListObject preserving formatting Aspose.Cells .NET example | clone Excel table to new worksheet using Aspose.Cells API | how to copy table range and recreate ListObject in Aspose.Cells | C# Aspose.Cells retain cell styles when copying tables between worksheets
// Tags: duplicate ListObject across worksheets Aspose.Cells | preserve formulas when cloning Excel tables C# | retain cell styles during table copy Aspose.Cells | create ListObject from range Aspose.Cells | clone Excel table with formatting Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example loads a workbook, extracts the first ListObject from Sheet1, copies each cell's value, formula, and style to a newly added worksheet, creates a new ListObject on the destination sheet covering the copied range, and saves the workbook as output.xlsx.
class DuplicateTableExample
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Source worksheet containing the table to duplicate
            Worksheet srcSheet = workbook.Worksheets["Sheet1"]; // adjust name as needed
            if (srcSheet == null)
            {
                Console.WriteLine("Source worksheet \"Sheet1\" not found.");
                return;
            }

            // Ensure the source worksheet contains at least one table (ListObject)
            if (srcSheet.ListObjects.Count == 0)
            {
                Console.WriteLine("No tables found on the source worksheet.");
                return;
            }

            // First table on the source sheet is the one to copy
            ListObject srcTable = srcSheet.ListObjects[0];

            // Destination worksheet (create a new one)
            Worksheet destSheet = workbook.Worksheets.Add("TableCopy");

            // Determine the source range of the table
            int srcStartRow = srcTable.DataRange.FirstRow;
            int srcStartCol = srcTable.DataRange.FirstColumn;
            int srcEndRow = srcStartRow + srcTable.DataRange.RowCount - 1;
            int srcEndCol = srcStartCol + srcTable.DataRange.ColumnCount - 1;

            CellArea srcArea = new CellArea
            {
                StartRow = srcStartRow,
                StartColumn = srcStartCol,
                EndRow = srcEndRow,
                EndColumn = srcEndCol
            };

            // Copy values, formulas and styles cell‑by‑cell
            for (int r = srcArea.StartRow; r <= srcArea.EndRow; r++)
            {
                for (int c = srcArea.StartColumn; c <= srcArea.EndColumn; c++)
                {
                    Cell srcCell = srcSheet.Cells[r, c];
                    Cell destCell = destSheet.Cells[r - srcArea.StartRow, c - srcArea.StartColumn];

                    // Copy value
                    destCell.PutValue(srcCell.Value);

                    // Copy formula (if any)
                    if (!string.IsNullOrEmpty(srcCell.Formula))
                        destCell.Formula = srcCell.Formula;

                    // Copy style
                    destCell.SetStyle(srcCell.GetStyle());
                }
            }

            // Destination range based on the size of the source table
            CellArea destArea = new CellArea
            {
                StartRow = 0,
                StartColumn = 0,
                EndRow = srcArea.EndRow - srcArea.StartRow,
                EndColumn = srcArea.EndColumn - srcArea.StartColumn
            };

            // Convert CellArea to an address string (e.g., "A1:C5")
            string startAddress = CellsHelper.CellIndexToName(destArea.StartRow, destArea.StartColumn);
            string endAddress = CellsHelper.CellIndexToName(destArea.EndRow, destArea.EndColumn);
            string destRange = $"{startAddress}:{endAddress}";

            // Add a new ListObject (table) on the destination sheet using the copied range
            // In older Aspose.Cells versions Add returns the index of the new table
            int destTableIndex = destSheet.ListObjects.Add("TableCopy", destRange, true);
            ListObject destTable = destSheet.ListObjects[destTableIndex];

            // (Optional) You can further customize destTable here if needed

            // Save the workbook with the duplicated table
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved as \"{outputPath}\".");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
