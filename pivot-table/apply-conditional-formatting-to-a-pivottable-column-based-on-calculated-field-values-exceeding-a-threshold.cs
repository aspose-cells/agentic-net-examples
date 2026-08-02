// Title: C# – Apply Conditional Formatting to a PivotTable Calculated Field with Aspose.Cells
// Description: Creates a workbook, adds product data, builds a PivotTable, defines a calculated field (Total = Price × Quantity), refreshes the pivot, and uses PivotConditionalFormat to highlight Total values above 100 with a LightCoral background and bold font.
// Keywords: Aspose.Cells | C# PivotTable | conditional formatting | calculated field | PivotConditionalFormat | highlight values | greater than threshold | Excel automation | .NET spreadsheet library
// Common Searches: Aspose.Cells conditional formatting for calculated field | C# highlight PivotTable values over 100 | add PivotConditionalFormat to data area Aspose.Cells | set background color for pivot calculated column | how to use calculated fields in Aspose.Cells PivotTable
// Developer Intent: Generate a PivotTable, add a calculated column, and apply visual cues to cells that exceed a defined numeric limit.
// Use Cases: Sales dashboards that flag totals above a budget limit. | Inventory reports that color‑code high‑value items for quick review. | Financial summaries that emphasize profit figures surpassing target thresholds.
// AI Prompts: Write C# code with Aspose.Cells to add a calculated field to a PivotTable and apply a LightCoral conditional format for values greater than 100. | Show how to set PivotConditionalFormat scope to the data area and define a GreaterThan condition for a pivot calculated column. | Explain the steps to refresh a PivotTable after adding a calculated field before applying conditional formatting in Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotConditionalFormatting
{
    // Creates a workbook, adds product data, builds a PivotTable, defines a calculated field (Total = Price × Quantity), refreshes the pivot, and uses PivotConditionalFormat to highlight Total values above 100 with a LightCoral background and bold font.
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
                // Columns: Product, Price, Quantity
                cells["A1"].Value = "Product";
                cells["B1"].Value = "Price";
                cells["C1"].Value = "Quantity";

                cells["A2"].Value = "Apple";
                cells["B2"].Value = 10;
                cells["C2"].Value = 5;

                cells["A3"].Value = "Banana";
                cells["B3"].Value = 8;
                cells["C3"].Value = 12;

                cells["A4"].Value = "Cherry";
                cells["B4"].Value = 15;
                cells["C4"].Value = 7;

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:C4", "E3", "SalesPivot");
                PivotTable pivot = sheet.PivotTables[pivotIndex];

                // Configure pivot fields
                pivot.AddFieldToArea(PivotFieldType.Row, "Product");          // Row field
                pivot.AddFieldToArea(PivotFieldType.Data, "Price");          // Data field 1
                pivot.AddFieldToArea(PivotFieldType.Data, "Quantity");       // Data field 2

                // Add a calculated field: Total = Price * Quantity
                pivot.AddCalculatedField("Total", "=Price*Quantity", true);

                // Refresh and calculate the pivot table to populate data
                pivot.RefreshData();
                pivot.CalculateData();

                // Add conditional formatting to the calculated field column (Total)
                int formatIdx = pivot.ConditionalFormats.Add();
                PivotConditionalFormat pcf = pivot.ConditionalFormats[formatIdx];

                // Apply the format to the calculated field (Total) in the data area
                pcf.AddFieldArea(PivotFieldType.Data, "Total");
                pcf.ScopeType = PivotConditionFormatScopeType.Data; // Apply to data fields

                // Define the condition: values greater than the threshold (e.g., 100)
                int conditionIdx = pcf.FormatConditions.AddCondition(FormatConditionType.CellValue);
                FormatCondition condition = pcf.FormatConditions[conditionIdx];
                condition.Operator = OperatorType.GreaterThan; // Correct enum value
                condition.Formula1 = "100";

                // Set the style to highlight cells exceeding the threshold
                condition.Style.BackgroundColor = Color.LightCoral;
                condition.Style.Font.IsBold = true;

                // Save the workbook
                string outputPath = "PivotTable_CalculatedField_ConditionalFormatting.xlsx";
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
