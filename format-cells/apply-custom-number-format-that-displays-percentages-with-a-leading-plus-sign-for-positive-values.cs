// Title: C# – Aspose.Cells: Custom Number Format to Show a Plus Sign for Positive Percentages
// Description: Demonstrates how to create a workbook, write percentage values, and apply a custom number format (+0.00%;-0.00%;0.00%) that prefixes positive percentages with a plus sign, negatives with a minus sign, and formats zero without a sign, then saves the file.
// Keywords: Aspose.Cells custom number format | C# percentage format plus sign | Excel positive negative zero format .NET | Aspose.Cells style custom format | display + sign for positive percentages | format cells as +0.00% -0.00% 0.00% | Aspose.Cells number format string | C# Excel formatting plus minus
// Common Searches: Aspose.Cells add plus sign to positive percentages | C# custom number format for positive negative zero values | How to format Excel cells as +0.00% in Aspose.Cells | Aspose.Cells percentage format with sign symbols | C# Aspose.Cells custom number format examples
// Developer Intent: Create a style with a custom number format that displays a leading '+' for positive percentages, '-' for negatives, and a neutral format for zero, then apply it to cells.
// Use Cases: Financial reports where gains must be marked with '+' and losses with '-' for instant visual cue. | KPI dashboards that require explicit sign indicators on percentage metrics. | Audit‑ready Excel exports that enforce a standardized sign‑aware percentage format.
// AI Prompts: Generate C# code using Aspose.Cells to format percentages with + for positives, - for negatives, and no sign for zero. | Show how to apply a custom number format to a range of cells in Aspose.Cells, handling zero values correctly. | Explain how to modify the format to show one decimal place (e.g., +0.0%;-0.0%;0.0%) while keeping the sign symbols.

using System;
using Aspose.Cells;

// Demonstrates how to create a workbook, write percentage values, and apply a custom number format (+0.00%;-0.00%;0.00%) that prefixes positive percentages with a plus sign, negatives with a minus sign, and formats zero without a sign, then saves the file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Sample values: positive, negative and zero percentages
        worksheet.Cells["A1"].PutValue(0.25);   // 25%
        worksheet.Cells["A2"].PutValue(-0.10); // -10%
        worksheet.Cells["A3"].PutValue(0);     // 0%

        // Create a style with a custom number format that shows a leading plus sign for positives
        Style percentStyle = workbook.CreateStyle();
        // Format: +0.00% for positive, -0.00% for negative, 0.00% for zero
        percentStyle.Custom = "+0.00%;-0.00%;0.00%";

        // Apply the style to the cells
        worksheet.Cells["A1"].SetStyle(percentStyle);
        worksheet.Cells["A2"].SetStyle(percentStyle);
        worksheet.Cells["A3"].SetStyle(percentStyle);

        // Save the workbook
        workbook.Save("PercentageWithPlusSign.xlsx");
    }
}
