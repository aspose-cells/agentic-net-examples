// Title: C# – Apply Conditional Formatting to a PivotTable Calculated Field with Aspose.Cells
// Description: Creates a workbook, builds a PivotTable from product data, adds a calculated field (Total = Price × Quantity), and uses PivotConditionalFormat to highlight Total values greater than 500 with a light‑salmon background and bold font, then saves the file.
// Keywords: Aspose.Cells | .NET | C# | PivotTable | Conditional Formatting | Calculated Field | PivotConditionalFormat | Excel automation | highlight values > 500 | Aspose.Cells example
// Common Searches: Aspose.Cells conditional formatting for PivotTable calculated field | C# highlight PivotTable column values above a threshold | How to use PivotConditionalFormat in Aspose.Cells | Set background color for PivotTable data area .NET | Apply conditional formatting to calculated field in Excel using Aspose
// Developer Intent: I want to automatically format cells in a PivotTable calculated column when their values exceed a specific limit using Aspose.Cells for .NET.
// Use Cases: Sales dashboard that flags products with total revenue over a target amount. | Financial report that draws attention to outlier totals for quick analysis. | Automated Excel export where high‑value rows are highlighted for reviewers.
// AI Prompts: Generate C# code with Aspose.Cells that adds a calculated field to a PivotTable and highlights values greater than 1000. | Show how to use PivotConditionalFormat to change font style and background color for a specific data field based on a numeric condition. | Provide an example that refreshes and calculates PivotTable data so conditional formatting on a calculated column takes effect.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotConditionalFormatting
{
    // Creates a workbook, builds a PivotTable from product data, adds a calculated field (Total = Price × Quantity), and uses PivotConditionalFormat to highlight Total values greater than 500 with a light‑salmon background and bold font, then saves the file.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data
                cells["A1"].Value = "Product";
                cells["B1"].Value = "Price";
                cells["C1"].Value = "Quantity";

                cells["A2"].Value = "Apple";
                cells["B2"].Value = 10;
                cells["C2"].Value = 30;

                cells["A3"].Value = "Banana";
                cells["B3"].Value = 8;
                cells["C3"].Value = 50;

                cells["A4"].Value = "Cherry";
                cells["B4"].Value = 12;
                cells["C4"].Value = 20;

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:C4", "E3", "SalesPivot");
                PivotTable pivot = sheet.PivotTables[pivotIndex];

                // Configure pivot fields
                pivot.AddFieldToArea(PivotFieldType.Row, "Product");
                pivot.AddFieldToArea(PivotFieldType.Data, "Price");
                pivot.AddFieldToArea(PivotFieldType.Data, "Quantity");

                // Add a calculated field: Total = Price * Quantity
                pivot.AddCalculatedField("Total", "=Price*Quantity", true);

                // Refresh the pivot cache so that the calculated field appears
                pivot.RefreshData();

                // Add conditional formatting to the calculated field column
                int formatIdx = pivot.ConditionalFormats.Add();
                PivotConditionalFormat pcf = pivot.ConditionalFormats[formatIdx];
                pcf.ScopeType = PivotConditionFormatScopeType.Data;
                pcf.AddFieldArea(PivotFieldType.Data, "Total");

                // Define the condition: values greater than 500
                int conditionIdx = pcf.FormatConditions.AddCondition(FormatConditionType.CellValue);
                FormatCondition condition = pcf.FormatConditions[conditionIdx];
                condition.Operator = OperatorType.GreaterThan;
                condition.Formula1 = "500";

                // Set visual style for cells that meet the condition
                condition.Style.BackgroundColor = Color.LightSalmon;
                condition.Style.Font.IsBold = true;

                // Calculate the pivot table data so that conditional formatting is applied
                pivot.CalculateData();

                // Save the workbook
                string outputPath = "PivotTable_With_CalculatedField_ConditionalFormatting.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
