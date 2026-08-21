// Title: C# – Merge V8:W9, apply red‑background conditional formatting for negatives, and export to HTML using Aspose.Cells
// Description: Load an Excel workbook, merge the range V8:W9, add a conditional formatting rule that highlights values < 0 with a red background, and save the worksheet as HTML while preserving merged areas via Aspose.Cells for .NET.
// Keywords: Aspose.Cells merge cells | C# conditional formatting negative values | export Excel to HTML Aspose.Cells | HtmlSaveOptions MergeAreas | CellArea V8 W9 | Aspose.Cells red background style | C# Excel to HTML conversion | Aspose.Cells .NET example | conditional formatting merged cells | Aspose.Cells GitHub
// Common Searches: Aspose.Cells merge V8 W9 C# | conditional formatting negative numbers Aspose.Cells .NET | export merged cells to HTML with Aspose.Cells | HtmlSaveOptions MergeAreas example | C# code to highlight negative values in Excel | Aspose.Cells tutorial for HTML export
// Developer Intent: Merge cells V8:W9, highlight negative numbers with a red background, and save the sheet as an HTML file.
// Use Cases: Financial dashboards where a total label spans V8:W9 and any negative amounts need visual emphasis before publishing online. | HTML‑based invoices that require merged header cells and automatic red‑highlighting of discount values below zero. | Web reports that combine merged title cells with conditional formatting to improve data readability.
// AI Prompts: Generate C# code with Aspose.Cells to merge V8:W9, apply a red‑background rule for values less than zero, and export the workbook to HTML preserving merged areas. | Explain the impact of HtmlSaveOptions.MergeAreas on the appearance of merged cells and conditional formatting in the HTML output. | Provide a step‑by‑step guide to add multiple conditional formatting rules to a merged cell block and export the result to HTML using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using System.Drawing;

namespace AsposeCellsExample
{
    // Load an Excel workbook, merge the range V8:W9, add a conditional formatting rule that highlights values < 0 with a red background, and save the worksheet as HTML while preserving merged areas via Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Merge cells V8:W9 (zero‑based indices: V=21, row 8=7)
            // totalRows = 2 (rows 8 and 9), totalColumns = 2 (V and W)
            cells.Merge(firstRow: 7, firstColumn: 21, totalRows: 2, totalColumns: 2);

            // Apply conditional formatting to highlight negative values in the merged area
            // 1. Add a new ConditionalFormatting entry
            int cfIndex = worksheet.ConditionalFormattings.Add();
            FormatConditionCollection fcc = worksheet.ConditionalFormattings[cfIndex];

            // 2. Define the range V8:W9
            CellArea area = new CellArea
            {
                StartRow = 7,
                StartColumn = 21,
                EndRow = 8,
                EndColumn = 22
            };
            fcc.AddArea(area);

            // 3. Add a condition: cell value less than 0
            int conditionIndex = fcc.AddCondition(
                type: FormatConditionType.CellValue,
                operatorType: OperatorType.LessThan,
                formula1: "0",
                formula2: null);

            // 4. Create a style for the condition (red background)
            Style negativeStyle = workbook.CreateStyle();
            negativeStyle.ForegroundColor = Color.Red;
            negativeStyle.Pattern = BackgroundType.Solid;

            // 5. Assign the style to the condition
            FormatCondition condition = fcc[conditionIndex];
            condition.Style = negativeStyle;

            // Export the workbook to HTML
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Ensure merged areas (including conditional formatting) are considered
                MergeAreas = true
            };
            workbook.Save("output.html", htmlOptions);
        }
    }
}
