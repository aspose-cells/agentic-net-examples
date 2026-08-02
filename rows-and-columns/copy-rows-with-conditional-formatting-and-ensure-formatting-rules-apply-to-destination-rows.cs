// Title: Copy rows with conditional formatting using Aspose.Cells for .NET (C#)
// Description: This example creates a source workbook, adds a conditional formatting rule that highlights values greater than 20 in column B, and copies the first five rows to a new workbook. The copy operation uses Cells.CopyRows with the ExtendToAdjacentRange option, and the ConditionalFormattings collection is transferred to retain the original formatting rules. Both workbooks are then saved.
// Keywords: Aspose.Cells copy rows C# | conditional formatting copy Aspose.Cells | Cells.CopyRows ExtendToAdjacentRange | ConditionalFormattings.Copy .NET | preserve formatting when copying rows | Aspose.Cells row duplication
// Common Searches: copy rows with conditional formatting Aspose.Cells .NET | how to keep conditional formatting after copying rows | CopyOptions ExtendToAdjacentRange example | transfer ConditionalFormattings between worksheets
// Developer Intent: Duplicate selected rows from one worksheet to another while automatically preserving any conditional formatting applied to the source range.
// Use Cases: Generate a reporting workbook that mirrors a styled data table from a master sheet. | Create department‑specific sheets by copying template rows with highlight rules intact. | Build a summary sheet that aggregates rows from multiple sources without losing conditional formatting.
// AI Prompts: Write C# code that copies rows 0‑10 from a source worksheet to a destination worksheet with Aspose.Cells and ensures all conditional formatting rules are retained. | Explain how the ExtendToAdjacentRange option influences the copying of merged cells and conditional formatting in Cells.CopyRows. | Provide a step‑by‑step tutorial for copying a ConditionalFormattings collection after using Cells.CopyRows in Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsCopyRowsWithConditionalFormatting
{
    // This example creates a source workbook, adds a conditional formatting rule that highlights values greater than 20 in column B, and copies the first five rows to a new workbook. The copy operation uses Cells.CopyRows with the ExtendToAdjacentRange option, and the ConditionalFormattings collection is transferred to retain the original formatting rules. Both workbooks are then saved.
    class Program
    {
        static void Main()
        {
            // ---------- Create source workbook and add data ----------
            Workbook sourceWorkbook = new Workbook();
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

            // Fill some data in rows 0-4
            for (int row = 0; row < 5; row++)
            {
                sourceSheet.Cells[row, 0].PutValue($"Item {row + 1}");
                sourceSheet.Cells[row, 1].PutValue(row * 10);
            }

            // Add a conditional formatting rule: highlight values > 20 in column B
            int cfIndex = sourceSheet.ConditionalFormattings.Add();
            FormatConditionCollection fcc = sourceSheet.ConditionalFormattings[cfIndex];
            fcc.AddCondition(FormatConditionType.CellValue, OperatorType.GreaterThan, "20", null);
            Style cfStyle = sourceWorkbook.CreateStyle();
            cfStyle.ForegroundColor = System.Drawing.Color.Yellow;
            cfStyle.Pattern = BackgroundType.Solid;
            fcc[0].Style = cfStyle;

            // Apply the conditional formatting to the range B1:B5 (rows 0-4, column 1)
            CellArea area = new CellArea
            {
                StartRow = 0,
                EndRow = 4,
                StartColumn = 1,
                EndColumn = 1
            };
            fcc.AddArea(area);

            // ---------- Create destination workbook ----------
            Workbook destWorkbook = new Workbook();
            Worksheet destSheet = destWorkbook.Worksheets[0];

            // ---------- Copy rows (data + formats) ----------
            // Use CopyOptions to extend ranges when copying (helps with merged cells, etc.)
            CopyOptions copyOptions = new CopyOptions
            {
                ExtendToAdjacentRange = true
            };
            // Copy the first 5 rows from source to destination starting at row 0
            destSheet.Cells.CopyRows(sourceSheet.Cells, 0, 0, 5, copyOptions);

            // ---------- Copy conditional formatting collection ----------
            // This ensures the conditional formatting rules are also present in the destination sheet
            destSheet.ConditionalFormattings.Copy(sourceSheet.ConditionalFormattings);

            // ---------- Save workbooks ----------
            sourceWorkbook.Save("SourceWithConditionalFormatting.xlsx");
            destWorkbook.Save("DestinationCopiedRows.xlsx");
        }
    }
}
