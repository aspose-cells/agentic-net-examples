// Title: Use Aspose.Cells for .NET to apply formula‑based conditional formatting that highlights sales values over 100 in an Excel worksheet
// Common Searches: aspnet apply conditional formatting with a formula using Aspose.Cells | c# Aspose.Cells highlight cells where value is greater than 100 | how to set background color for cells based on a formula in Aspose.Cells .NET | example of conditional formatting range B2:B6 with Aspose.Cells | apply bold white font to cells that exceed a threshold in Aspose.Cells
// Tags: formula conditional formatting Aspose.Cells | highlight cells over threshold .NET | apply solid background style Aspose.Cells | conditional formatting range B2:B6 Excel | bold white font on condition Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;

// The sample creates a workbook, fills column A with item names and column B with sales figures, then adds a formula‑based conditional formatting rule to the range B2:B6 using the expression "$B2>100". Cells meeting the condition are styled with a LightCoral solid background, white bold font, and the workbook is saved as ConditionalFormatting_Threshold.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data (e.g., sales figures)
            cells["A1"].PutValue("Item");
            cells["B1"].PutValue("Sales");
            string[] items = { "Apple", "Banana", "Cherry", "Date", "Elderberry" };
            double[] sales = { 85.0, 120.5, 95.0, 150.0, 70.0 };

            for (int i = 0; i < items.Length; i++)
            {
                cells[i + 2, 0].PutValue(items[i]);   // Column A
                cells[i + 2, 1].PutValue(sales[i]);   // Column B
            }

            // Define the range to which the conditional formatting will be applied (B2:B6)
            int startRow = 1;                 // zero‑based index (row 2)
            int endRow = items.Length;        // last data row (zero‑based)
            int startColumn = 1;              // column B
            int endColumn = 1;                // column B

            // Add a conditional formatting collection to the worksheet
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection fcs = sheet.ConditionalFormattings[cfIndex];

            // Define the area for the conditional formatting
            CellArea area = new CellArea
            {
                StartRow = startRow,
                EndRow = endRow,
                StartColumn = startColumn,
                EndColumn = endColumn
            };
            fcs.AddArea(area);

            // Add a condition of type Expression (formula based)
            int conditionIndex = fcs.AddCondition(FormatConditionType.Expression);
            FormatCondition condition = fcs[conditionIndex];

            // Formula: highlight cells where the sales value exceeds 100
            condition.Formula1 = "$B2>100";

            // Define the style to apply when the condition is true
            Style style = condition.Style;
            style.ForegroundColor = Color.LightCoral;
            style.Pattern = BackgroundType.Solid;
            style.Font.Color = Color.White;
            style.Font.IsBold = true;

            // Recalculate formulas (optional)
            workbook.CalculateFormula();

            // Save the workbook
            string outputPath = "ConditionalFormatting_Threshold.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred:");
            Console.WriteLine(ex.Message);
        }
    }
}
