using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Settings;

namespace AsposeCellsPivotGlobalizationDemo
{
    // Custom globalization settings that override various pivot table labels
    public class CustomPivotGlobalizationSettings : SettablePivotGlobalizationSettings
    {
        public CustomPivotGlobalizationSettings()
        {
            // Set custom texts for pivot table UI elements
            SetTextOfColumnLabels("Custom Column Headers");
            SetTextOfRowLabels("Custom Row Headers");
            SetTextOfTotal("Custom Total");
            SetTextOfGrandTotal("Custom Grand Total");
            SetTextOfAll("Custom All");
            SetTextOfDataFieldHeader("Custom Data Header");
            SetTextOfEmptyData("Custom Empty");
            SetTextOfMultipleItems("Custom Multiple Items");
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet dataSheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
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

            // Add a new worksheet for the pivot table
            Worksheet pivotSheet = workbook.Worksheets.Add("PivotTable");

            // Create the pivot table
            int pivotIndex = pivotSheet.PivotTables.Add("A1:C4", "E5", "SalesPivot");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

            // Configure pivot fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);    // Product
            pivotTable.AddFieldToArea(PivotFieldType.Column, 1); // Region
            pivotTable.AddFieldToArea(PivotFieldType.Data, 2);   // Sales

            // Apply custom globalization settings
            GlobalizationSettings globalSettings = new GlobalizationSettings();
            globalSettings.PivotSettings = new CustomPivotGlobalizationSettings();
            workbook.Settings.GlobalizationSettings = globalSettings;

            // Refresh and calculate the pivot table to apply the custom labels
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Verify that the custom texts are applied
            var pivotGlobals = (CustomPivotGlobalizationSettings)workbook.Settings.GlobalizationSettings.PivotSettings;

            Console.WriteLine("Verification of custom globalization labels:");
            Console.WriteLine($"Column Labels: {pivotGlobals.GetTextOfColumnLabels()}");
            Console.WriteLine($"Row Labels: {pivotGlobals.GetTextOfRowLabels()}");
            Console.WriteLine($"Total: {pivotGlobals.GetTextOfTotal()}");
            Console.WriteLine($"Grand Total: {pivotGlobals.GetTextOfGrandTotal()}");
            Console.WriteLine($"(All): {pivotGlobals.GetTextOfAll()}");
            Console.WriteLine($"Data Field Header: {pivotGlobals.GetTextOfDataFieldHeader()}");
            Console.WriteLine($"Empty Data: {pivotGlobals.GetTextOfEmptyData()}");
            Console.WriteLine($"Multiple Items: {pivotGlobals.GetTextOfMultipleItems()}");

            // Save the workbook
            workbook.Save("CustomPivotGlobalization.xlsx");
            Console.WriteLine("Workbook saved as 'CustomPivotGlobalization.xlsx'.");
        }
    }
}