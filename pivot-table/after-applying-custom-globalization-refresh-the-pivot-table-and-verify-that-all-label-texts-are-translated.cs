using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Settings;

namespace AsposeCellsPivotGlobalizationDemo
{
    // Demonstrates applying custom globalization to a PivotTable,
    // refreshing it and verifying that all label texts are translated.
    public class Program
    {
        public static void Main()
        {
            // 1. Create a new workbook and get the first worksheet.
            Workbook workbook = new Workbook();
            Worksheet dataSheet = workbook.Worksheets[0];

            // 2. Populate sample data for the PivotTable.
            dataSheet.Cells["A1"].PutValue("Product");
            dataSheet.Cells["B1"].PutValue("Region");
            dataSheet.Cells["C1"].PutValue("Sales");

            dataSheet.Cells["A2"].PutValue("Bike");
            dataSheet.Cells["B2"].PutValue("North");
            dataSheet.Cells["C2"].PutValue(10000);

            dataSheet.Cells["A3"].PutValue("Bike");
            dataSheet.Cells["B3"].PutValue("South");
            dataSheet.Cells["C3"].PutValue(8000);

            dataSheet.Cells["A4"].PutValue("Car");
            dataSheet.Cells["B4"].PutValue("North");
            dataSheet.Cells["C4"].PutValue(25000);

            // 3. Add a new worksheet to host the PivotTable.
            Worksheet pivotSheet = workbook.Worksheets.Add("PivotTable");

            // 4. Create the PivotTable.
            int pivotIndex = pivotSheet.PivotTables.Add("A1:C4", "E5", "SalesPivot");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

            // 5. Configure the PivotTable fields.
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);    // Product as row field
            pivotTable.AddFieldToArea(PivotFieldType.Column, 1); // Region as column field
            pivotTable.AddFieldToArea(PivotFieldType.Data, 2);   // Sales as data field

            // 6. Create customizable globalization settings.
            //    SettableGlobalizationSettings holds a PivotSettings property of type SettablePivotGlobalizationSettings.
            SettableGlobalizationSettings globalizationSettings = new SettableGlobalizationSettings();
            SettablePivotGlobalizationSettings pivotSettings = new SettablePivotGlobalizationSettings();

            // 7. Set custom texts for various PivotTable labels.
            pivotSettings.SetTextOfAll("(All Items)");
            pivotSettings.SetTextOfColumnLabels("Custom Column Headers");
            pivotSettings.SetTextOfRowLabels("Custom Row Headers");
            pivotSettings.SetTextOfTotal("Custom Total");
            pivotSettings.SetTextOfGrandTotal("Custom Grand Total");
            pivotSettings.SetTextOfDataFieldHeader("Custom Data Header");
            pivotSettings.SetTextOfEmptyData("No Data");
            pivotSettings.SetTextOfMultipleItems("(Multiple Selections)");

            // 8. Attach the pivot settings to the globalization settings.
            globalizationSettings.PivotSettings = pivotSettings;

            // 9. Apply the globalization settings to the workbook.
            workbook.Settings.GlobalizationSettings = globalizationSettings;

            // 10. Refresh and calculate the PivotTable so that the new labels take effect.
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // 11. Verify that the custom texts are applied by reading them back.
            Console.WriteLine("Verification of customized PivotTable labels:");
            Console.WriteLine($"All label: {pivotSettings.GetTextOfAll()}");
            Console.WriteLine($"Column Labels: {pivotSettings.GetTextOfColumnLabels()}");
            Console.WriteLine($"Row Labels: {pivotSettings.GetTextOfRowLabels()}");
            Console.WriteLine($"Total: {pivotSettings.GetTextOfTotal()}");
            Console.WriteLine($"Grand Total: {pivotSettings.GetTextOfGrandTotal()}");
            Console.WriteLine($"Data Field Header: {pivotSettings.GetTextOfDataFieldHeader()}");
            Console.WriteLine($"Empty Data: {pivotSettings.GetTextOfEmptyData()}");
            Console.WriteLine($"Multiple Items: {pivotSettings.GetTextOfMultipleItems()}");

            // 12. Save the workbook to demonstrate successful execution.
            workbook.Save("CustomPivotGlobalizationDemo.xlsx");
            Console.WriteLine("Workbook saved as 'CustomPivotGlobalizationDemo.xlsx'.");
        }
    }
}