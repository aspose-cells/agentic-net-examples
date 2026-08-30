// Title: How to apply conditional formatting to a calculated "Total" field in an Aspose.Cells PivotTable using C#
// AI Prompts: Write C# code that creates a PivotTable, adds a calculated field named Total (Price * Quantity), and applies a PivotConditionalFormat to highlight cells where Total exceeds 1000. | Show how to set PivotConditionFormatScopeType to Data and specify the field area for a calculated data field in Aspose.Cells. | Demonstrate configuring a FormatCondition with OperatorType.GreaterThan and a custom background color for a PivotTable calculated column.
// Common Searches: Aspose.Cells C# conditional formatting on pivot table calculated field Total | Highlight pivot table column values greater than 1000 using Aspose.Cells | Set PivotConditionalFormat scope to data field in C# Aspose.Cells example | Add calculated field to PivotTable and apply cell background color in Aspose.Cells | How to use FormatCondition with OperatorType.GreaterThan in Aspose.Cells pivot
// Tags: aspocells pivotconditionalformat calculated field | c# conditional formatting pivot table aspocells | highlight values over threshold aspocells | add calculated field to aspocells pivot | formatcondition greaterthan aspocells

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using System.Drawing;

namespace AsposeCellsPivotConditionalFormatting
{
    // The example creates a workbook, populates product data, builds a PivotTable with Product rows and Price/Quantity data fields, adds a calculated field "Total" (Price × Quantity), refreshes the pivot, and then applies a PivotConditionalFormat scoped to data fields. The format targets the "Total" column, uses a GreaterThan condition with a threshold of 1000, and sets a light coral background before saving the file.
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
                cells["C2"].Value = 120;

                cells["A3"].Value = "Banana";
                cells["B3"].Value = 8;
                cells["C3"].Value = 150;

                cells["A4"].Value = "Cherry";
                cells["B4"].Value = 12;
                cells["C4"].Value = 90;

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:C4", "E3", "SalesPivot");
                PivotTable pivot = sheet.PivotTables[pivotIndex];

                // Configure pivot fields
                pivot.AddFieldToArea(PivotFieldType.Row, "Product");          // Row field
                pivot.AddFieldToArea(PivotFieldType.Data, "Price");          // Data field
                pivot.AddFieldToArea(PivotFieldType.Data, "Quantity");       // Data field

                // Add a calculated field: Total = Price * Quantity, and drag it to the data area
                pivot.AddCalculatedField("Total", "=Price*Quantity", true);

                // Refresh and calculate the pivot table to populate data
                pivot.RefreshData();   // Correct API call
                pivot.CalculateData();

                // Add conditional formatting targeting the calculated field column
                int formatIdx = pivot.ConditionalFormats.Add();
                PivotConditionalFormat pcf = pivot.ConditionalFormats[formatIdx];

                // Apply to data fields (the calculated field is a data field)
                pcf.ScopeType = PivotConditionFormatScopeType.Data;

                // Define the area as the calculated field "Total"
                pcf.AddFieldArea(PivotFieldType.Data, "Total");

                // Create a condition: values greater than 1000
                int conditionIdx = pcf.FormatConditions.AddCondition(FormatConditionType.CellValue);
                FormatCondition condition = pcf.FormatConditions[conditionIdx];
                condition.Operator = OperatorType.GreaterThan;
                condition.Formula1 = "1000";

                // Set the formatting style (e.g., light red background)
                condition.Style.BackgroundColor = Color.LightCoral;

                // Save the workbook
                workbook.Save("PivotTable_With_CalculatedField_ConditionalFormatting.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
