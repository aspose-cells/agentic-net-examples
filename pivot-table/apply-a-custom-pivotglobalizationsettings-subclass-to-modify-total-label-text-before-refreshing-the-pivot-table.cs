using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Settings;

namespace AsposeCellsExamples
{
    // Custom globalization settings that changes the "Total" label text
    public class CustomPivotGlobalizationSettings : PivotGlobalizationSettings
    {
        // Override the method that provides the text for the "Total" label
        public override string GetTextOfTotal()
        {
            // Return the custom label you want to appear in the pivot table
            return "Custom Total Label";
        }
    }

    public class ApplyCustomPivotGlobalization
    {
        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet dataSheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            dataSheet.Cells["A1"].PutValue("Category");
            dataSheet.Cells["B1"].PutValue("Amount");
            dataSheet.Cells["A2"].PutValue("Fruit");
            dataSheet.Cells["B2"].PutValue(120);
            dataSheet.Cells["A3"].PutValue("Fruit");
            dataSheet.Cells["B3"].PutValue(80);
            dataSheet.Cells["A4"].PutValue("Vegetable");
            dataSheet.Cells["B4"].PutValue(150);
            dataSheet.Cells["A5"].PutValue("Vegetable");
            dataSheet.Cells["B5"].PutValue(200);

            // Add a new worksheet that will contain the pivot table
            Worksheet pivotSheet = workbook.Worksheets.Add("PivotTable");

            // Create the pivot table on the pivot sheet
            int pivotIndex = pivotSheet.PivotTables.Add("A1:B5", "D1", "MyPivot");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

            // Configure the pivot table fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Category as row field
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1); // Amount as data field

            // Instantiate the custom globalization settings
            CustomPivotGlobalizationSettings customSettings = new CustomPivotGlobalizationSettings();

            // Assign the custom settings to the workbook's globalization settings
            workbook.Settings.GlobalizationSettings = new GlobalizationSettings();
            workbook.Settings.GlobalizationSettings.PivotSettings = customSettings;

            // Refresh and calculate the pivot table so that the custom label takes effect
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("CustomPivotTotalLabel.xlsx");
        }
    }

    // Entry point for demonstration
    class Program
    {
        static void Main()
        {
            ApplyCustomPivotGlobalization.Run();
        }
    }
}