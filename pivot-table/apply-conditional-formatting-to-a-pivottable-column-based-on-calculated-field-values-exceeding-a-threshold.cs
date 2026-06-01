using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotConditionalFormatting
{
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

                // -------------------------------------------------
                // Sample data for the pivot table
                // -------------------------------------------------
                // Header row
                cells["A1"].Value = "Region";
                cells["B1"].Value = "Product";
                cells["C1"].Value = "UnitsSold";
                cells["D1"].Value = "UnitPrice";

                // Data rows
                cells["A2"].Value = "North";   cells["B2"].Value = "Apple";   cells["C2"].Value = 120; cells["D2"].Value = 3.5;
                cells["A3"].Value = "North";   cells["B3"].Value = "Banana";  cells["C3"].Value = 80;  cells["D3"].Value = 2.0;
                cells["A4"].Value = "South";   cells["B4"].Value = "Apple";   cells["C4"].Value = 150; cells["D4"].Value = 3.5;
                cells["A5"].Value = "South";   cells["B5"].Value = "Banana";  cells["C5"].Value = 60;  cells["D5"].Value = 2.0;
                cells["A6"].Value = "East";    cells["B6"].Value = "Apple";   cells["C6"].Value = 200; cells["D6"].Value = 3.5;
                cells["A7"].Value = "East";    cells["B7"].Value = "Banana";  cells["C7"].Value = 90;  cells["D7"].Value = 2.0;
                cells["A8"].Value = "West";    cells["B8"].Value = "Apple";   cells["C8"].Value = 110; cells["D8"].Value = 3.5;
                cells["A9"].Value = "West";    cells["B9"].Value = "Banana";  cells["C9"].Value = 70;  cells["D9"].Value = 2.0;

                // -------------------------------------------------
                // Create the pivot table
                // -------------------------------------------------
                // Data source range: A1:D9
                // Destination top‑left cell: F3
                int pivotIndex = sheet.PivotTables.Add("A1:D9", "F3", "SalesPivot");
                PivotTable pivot = sheet.PivotTables[pivotIndex];

                // Add fields to the pivot table
                pivot.AddFieldToArea(PivotFieldType.Row, "Region");      // Row field
                pivot.AddFieldToArea(PivotFieldType.Column, "Product"); // Column field
                pivot.AddFieldToArea(PivotFieldType.Data, "UnitsSold"); // Data field (sum of units)

                // -------------------------------------------------
                // Add a calculated field: TotalRevenue = UnitsSold * UnitPrice
                // -------------------------------------------------
                pivot.AddCalculatedField("TotalRevenue", "=UnitsSold*UnitPrice", true);

                // -------------------------------------------------
                // Apply conditional formatting to the calculated field column
                // -------------------------------------------------
                // 1. Create a new conditional format entry
                int formatIdx = pivot.ConditionalFormats.Add();
                PivotConditionalFormat pcf = pivot.ConditionalFormats[formatIdx];

                // 2. Scope: apply to the data fields (the calculated field is now a data field)
                pcf.ScopeType = PivotConditionFormatScopeType.Data;

                // 3. Define the area: the calculated field "TotalRevenue"
                pcf.AddFieldArea(PivotFieldType.Data, "TotalRevenue");

                // 4. Add a format condition – highlight cells where revenue > 500
                int conditionIdx = pcf.FormatConditions.AddCondition(FormatConditionType.CellValue);
                FormatCondition condition = pcf.FormatConditions[conditionIdx];
                condition.Operator = OperatorType.GreaterThan; // Correct enum value
                condition.Formula1 = "500"; // Threshold

                // Style: light orange background, bold font
                condition.Style.BackgroundColor = Color.Orange;
                condition.Style.Font.IsBold = true;

                // -------------------------------------------------
                // Refresh and calculate the pivot table to apply changes
                // -------------------------------------------------
                pivot.RefreshData();
                pivot.CalculateData();

                // -------------------------------------------------
                // Save the workbook
                // -------------------------------------------------
                string outputPath = "PivotTable_With_CalculatedField_ConditionalFormatting.xlsx";

                // Ensure the directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                // Log the exception details
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}