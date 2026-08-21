// Title: Copy Rows with Conditional Formatting in Aspose.Cells for .NET (C#)
// Description: Demonstrates how to copy rows from a source worksheet to a destination worksheet using Aspose.Cells, while preserving conditional formatting rules. The example creates sample data, adds a red‑background rule for values between 10 and 30 in column A, copies rows with the ExtendToAdjacentRange option, recreates the conditional formatting on the target sheet with an adjusted CellArea, and saves the workbook as RowsWithConditionalFormatting.xlsx.
// Keywords: Aspose.Cells copy rows C# | conditional formatting copy Aspose | CopyRows ExtendToAdjacentRange | preserve conditional formatting .NET | adjust CellArea after copy | duplicate rows with styles Aspose.Cells | C# Excel conditional formatting transfer
// Common Searches: copy rows and keep conditional formatting Aspose.Cells | CopyRows ExtendToAdjacentRange example C# | how to transfer conditional formatting after copying rows | Aspose.Cells adjust conditional formatting range | duplicate formatted rows in Excel using Aspose
// Developer Intent: Copy rows between worksheets and ensure the original conditional formatting rules are applied to the new rows.
// Use Cases: Generate a report by duplicating a formatted template row into multiple result rows while retaining color‑coded thresholds. | Consolidate data from several source sheets into a summary sheet without losing conditional highlighting. | Automate the creation of a new workbook that mirrors a source sheet’s conditional formatting after row insertion or reordering.
// AI Prompts: Write C# code with Aspose.Cells that copies a range of rows and automatically updates any conditional formatting to the new row positions. | Show how to use CopyOptions.ExtendToAdjacentRange when copying rows and then replicate the source conditional formatting on the destination sheet. | Explain the steps to duplicate conditional formatting rules after copying rows in Aspose.Cells, including style copying and CellArea offset adjustments.

using System;
using Aspose.Cells;
using System.Drawing;

namespace AsposeCellsCopyRowsWithConditionalFormatting
{
    // Demonstrates how to copy rows from a source worksheet to a destination worksheet using Aspose.Cells, while preserving conditional formatting rules. The example creates sample data, adds a red‑background rule for values between 10 and 30 in column A, copies rows with the ExtendToAdjacentRange option, recreates the conditional formatting on the target sheet with an adjusted CellArea, and saves the workbook as RowsWithConditionalFormatting.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // ---------- Create source workbook and add data ----------
                Workbook sourceWorkbook = new Workbook();
                Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

                // Fill sample data in rows 0‑4, columns A‑B
                for (int row = 0; row < 5; row++)
                {
                    sourceSheet.Cells[row, 0].PutValue(row * 10);          // Column A
                    sourceSheet.Cells[row, 1].PutValue(row * 20 + 5);      // Column B
                }

                // ---------- Add conditional formatting ----------
                // Highlight cells in column A whose value is between 10 and 30
                int cfIndex = sourceSheet.ConditionalFormattings.Add();
                FormatConditionCollection srcFcc = sourceSheet.ConditionalFormattings[cfIndex];
                srcFcc.AddCondition(FormatConditionType.CellValue, OperatorType.Between, "10", "30");

                // Create a style (red background) for the condition
                Style cfStyle = sourceWorkbook.CreateStyle();
                cfStyle.BackgroundColor = Color.Red;
                srcFcc[0].Style = cfStyle;

                // Apply the conditional formatting to the range A1:A5
                CellArea srcArea = new CellArea
                {
                    StartRow = 0,
                    EndRow = 4,
                    StartColumn = 0,
                    EndColumn = 0
                };
                srcFcc.AddArea(srcArea);   // Add area to the collection, not to the condition

                // ---------- Create destination workbook ----------
                Workbook destWorkbook = new Workbook();
                Worksheet destSheet = destWorkbook.Worksheets[0];

                // ---------- Copy rows (including formats) ----------
                Cells sourceCells = sourceSheet.Cells;
                Cells destCells = destSheet.Cells;

                CopyOptions copyOptions = new CopyOptions
                {
                    // Extend formatting to adjacent ranges when copying rows
                    ExtendToAdjacentRange = true
                };

                // Copy rows 0‑4 to destination starting at row index 2
                destCells.CopyRows(sourceCells, 0, 2, 5, copyOptions);

                // ---------- Copy conditional formatting ----------
                // Recreate the conditional formatting in the destination sheet with adjusted area
                int destCfIndex = destSheet.ConditionalFormattings.Add();
                FormatConditionCollection destFcc = destSheet.ConditionalFormattings[destCfIndex];

                // Assume a single condition was added to the source sheet
                FormatCondition srcCondition = srcFcc[0];
                int destCondIdx = destFcc.AddCondition(
                    srcCondition.Type,
                    srcCondition.Operator,
                    srcCondition.Formula1,
                    srcCondition.Formula2);
                FormatCondition destCondition = destFcc[destCondIdx];

                // Copy the style
                destCondition.Style = srcCondition.Style;

                // Adjust the area to the destination offset (row 2)
                CellArea destArea = srcArea;
                destArea.StartRow += 2; // offset where rows were copied
                destArea.EndRow += 2;
                destFcc.AddArea(destArea);   // Add area to the destination collection

                // ---------- Save the result ----------
                destWorkbook.Save("RowsWithConditionalFormatting.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
