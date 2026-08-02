// Title: C# – Merge Cells with Aspose.Cells while Preserving Original Formatting
// Description: Demonstrates how to merge a range (A1:B2) in an Aspose.Cells workbook without altering existing cell styles. The example captures the style of the upper‑left cell, merges the range using Cells.Merge, optionally reapplies the captured style with SetStyle(explicitFlag:true), and saves the file as MergedPreserveStyle.xlsx.
// Keywords: Aspose.Cells merge cells preserve style | C# Aspose.Cells keep formatting after merge | Cells.Merge retain original style | SetStyle explicit flag Aspose.Cells | .NET workbook merge without losing formatting
// Common Searches: merge cells in Aspose.Cells without changing formatting | preserve cell style after merging range Aspose.Cells .NET | how to keep bold header style when merging cells Aspose | Aspose.Cells SetStyle after Cells.Merge | C# merge A1:B2 keep original formatting
// Developer Intent: Merge a cell range in an Aspose.Cells workbook while ensuring that all pre‑existing formatting (fonts, colors, patterns) remains unchanged.
// Use Cases: Combine header and sub‑header cells into a single merged cell without losing bold or italic styling. | Consolidate a data block (A1:B2) while preserving font size, color, and background patterns. | Capture a cell’s style before a merge operation and reapply it to guarantee no visual changes.
// AI Prompts: Write C# code using Aspose.Cells to merge a specified range and automatically retain all cell styles. | Show how to capture a cell’s style, perform Cells.Merge, and then reapply the style with SetStyle(explicitFlag:true) to avoid formatting loss. | Explain the interaction between Aspose.Cells’ Merge method and cell styles, and how the explicit flag in SetStyle preserves formatting.

using System;
using Aspose.Cells;

namespace AsposeCellsMergePreserveStyle
{
    // Demonstrates how to merge a range (A1:B2) in an Aspose.Cells workbook without altering existing cell styles. The example captures the style of the upper‑left cell, merges the range using Cells.Merge, optionally reapplies the captured style with SetStyle(explicitFlag:true), and saves the file as MergedPreserveStyle.xlsx.
    class Program
    {
        static void Main()
        {
            // ---------- Create a new workbook ----------
            Workbook workbook = new Workbook();                     // create
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // ---------- Populate cells with values and distinct styles ----------
            // Cell A1
            cells["A1"].PutValue("Header");
            Style styleA1 = workbook.CreateStyle();
            styleA1.Font.IsBold = true;
            styleA1.Font.Color = System.Drawing.Color.White;
            styleA1.ForegroundColor = System.Drawing.Color.DarkBlue;
            styleA1.Pattern = BackgroundType.Solid;
            cells["A1"].SetStyle(styleA1);

            // Cell B1
            cells["B1"].PutValue("SubHeader");
            Style styleB1 = workbook.CreateStyle();
            styleB1.Font.IsItalic = true;
            styleB1.ForegroundColor = System.Drawing.Color.LightGray;
            styleB1.Pattern = BackgroundType.Solid;
            cells["B1"].SetStyle(styleB1);

            // Cell A2 (will be part of the merged area)
            cells["A2"].PutValue("Data");
            Style styleA2 = workbook.CreateStyle();
            styleA2.Font.Size = 12;
            styleA2.Font.Color = System.Drawing.Color.Black;
            cells["A2"].SetStyle(styleA2);

            // ---------- Preserve original style before merging ----------
            // Capture the style of the upper‑left cell (A1) – this style will be
            // automatically retained after the merge because the Merge operation
            // does not modify cell styles.
            Style originalStyle = cells["A1"].GetStyle();

            // ---------- Merge the range A1:B2 ----------
            // Using Cells.Merge (firstRow, firstColumn, totalRows, totalColumns)
            // This combines the four cells into a single merged cell.
            cells.Merge(0, 0, 2, 2);                               // merge

            // ---------- Re‑apply the captured style explicitly (optional) ----------
            // If you want to guarantee that no style changes occurred, re‑apply
            // the original style with explicitFlag = true. This overwrites only
            // the properties that were explicitly set in the style, leaving all
            // other formatting untouched.
            cells["A1"].SetStyle(originalStyle, true);            // setstyle

            // ---------- Save the workbook ----------
            workbook.Save("MergedPreserveStyle.xlsx");             // save
        }
    }
}
