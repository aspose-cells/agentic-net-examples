// Title: Apply custom PivotTable globalization settings and refresh the table using Aspose.Cells for .NET (C#)
// AI Prompts: Implement a subclass of SettablePivotGlobalizationSettings that overrides column, row, total, grand total, and all-item label texts, then assign it to a workbook's GlobalizationSettings. | Refresh and calculate the PivotTable after applying the custom globalization to make the new labels visible. | Write code to output the overridden label values to the console and save the workbook with the customized PivotTable to an .xlsx file.
// Common Searches: Aspose.Cells C# set custom pivot table column label text | how to change pivot table row header language using globalization settings Aspose.Cells | refresh pivot table after modifying globalization settings in .NET | verify custom pivot table labels programmatically with Aspose.Cells
// Tags: SettablePivotGlobalizationSettings custom label overrides | Aspose.Cells pivot table refresh after globalization | C# pivot table label customization Aspose.Cells | globalization settings for pivot tables .NET | save workbook with customized pivot labels

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Settings;

namespace AsposeCellsPivotGlobalizationDemo
{
    // Custom globalization settings that modify various PivotTable label texts
    // Demonstrates creating a workbook, adding sample data, building a PivotTable, defining a CustomPivotGlobalizationSettings class that overrides column, row, total, grand total, and all-item labels, applying these settings via GlobalizationSettings, refreshing and calculating the PivotTable, printing the new label texts, and saving the file as CustomPivotGlobalizationDemo.xlsx.
    public class CustomPivotGlobalizationSettings : SettablePivotGlobalizationSettings
    {
        public CustomPivotGlobalizationSettings()
        {
            // Set custom texts for the PivotTable labels
            SetTextOfColumnLabels("Custom Column Headers");
            SetTextOfRowLabels("Custom Row Headers");
            SetTextOfTotal("Custom Total");
            SetTextOfGrandTotal("Custom Grand Total");
            SetTextOfAll("All Items");
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet dataSheet = workbook.Worksheets[0];

            // Populate sample data for the PivotTable
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

            // Create a PivotTable on the same sheet (you can also use a separate sheet)
            int pivotIndex = dataSheet.PivotTables.Add("A1:C4", "E5", "SalesPivot");
            PivotTable pivotTable = dataSheet.PivotTables[pivotIndex];

            // Configure the PivotTable fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);      // Product as Row field
            pivotTable.AddFieldToArea(PivotFieldType.Column, 1);   // Region as Column field
            pivotTable.AddFieldToArea(PivotFieldType.Data, 2);     // Sales as Data field

            // Apply custom globalization settings to the workbook
            var customPivotSettings = new CustomPivotGlobalizationSettings();
            GlobalizationSettings globalization = new GlobalizationSettings();
            globalization.PivotSettings = customPivotSettings;
            workbook.Settings.GlobalizationSettings = globalization;

            // Refresh and calculate the PivotTable to apply the custom labels
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Verify that the custom texts are set (output to console)
            Console.WriteLine("Verification of custom PivotTable label texts:");
            Console.WriteLine($"Column Labels: {customPivotSettings.GetTextOfColumnLabels()}");
            Console.WriteLine($"Row Labels: {customPivotSettings.GetTextOfRowLabels()}");
            Console.WriteLine($"Total: {customPivotSettings.GetTextOfTotal()}");
            Console.WriteLine($"Grand Total: {customPivotSettings.GetTextOfGrandTotal()}");
            Console.WriteLine($"(All) label: {customPivotSettings.GetTextOfAll()}");

            // Save the workbook
            workbook.Save("CustomPivotGlobalizationDemo.xlsx");
        }
    }
}
