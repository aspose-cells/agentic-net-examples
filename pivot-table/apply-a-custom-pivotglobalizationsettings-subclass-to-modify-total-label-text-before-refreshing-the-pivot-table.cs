// Title: How to override the Pivot table Total label in Aspose.Cells using a custom PivotGlobalizationSettings subclass (C#)
// AI Prompts: Create a C# class that inherits from PivotGlobalizationSettings, override GetTextOfTotal to return a custom string, assign the instance to workbook.Settings.GlobalizationSettings.PivotSettings, then refresh and calculate the pivot table. | Change the custom total caption in the example, update the output filename, and verify that the new label appears in the saved Excel workbook. | Show how to reuse the same custom PivotGlobalizationSettings object for multiple pivot tables in a workbook and refresh each table after applying the settings.
// Common Searches: asp.net change total caption in Aspose.Cells pivot table | c# Aspose.Cells set custom text for pivot total | example overriding GetTextOfTotal in Aspose.Cells | apply same globalization settings to several pivot tables Aspose.Cells | how to customize pivot table total text using Aspose.Cells C#
// Tags: custom total label via PivotGlobalizationSettings | Aspose.Cells pivot table globalization customization | C# override total caption in pivot table | apply custom globalization to multiple pivots | Excel pivot total text customization Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Settings;

namespace AsposeCellsExamples
{
    // The sample creates a workbook with sample data, adds a pivot table, defines a CustomPivotGlobalizationSettings class that overrides GetTextOfTotal to return "My Custom Total", assigns this custom settings object to the workbook's globalization settings, refreshes and calculates the pivot table, and saves the result as CustomTotalLabelDemo.xlsx.
    class CustomTotalLabelDemo
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet for data
            Workbook workbook = new Workbook();
            Worksheet dataSheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            dataSheet.Cells["A1"].PutValue("Category");
            dataSheet.Cells["B1"].PutValue("Amount");
            dataSheet.Cells["A2"].PutValue("Food");
            dataSheet.Cells["B2"].PutValue(120);
            dataSheet.Cells["A3"].PutValue("Drink");
            dataSheet.Cells["B3"].PutValue(80);
            dataSheet.Cells["A4"].PutValue("Food");
            dataSheet.Cells["B4"].PutValue(150);
            dataSheet.Cells["A5"].PutValue("Drink");
            dataSheet.Cells["B5"].PutValue(70);

            // Add a new worksheet that will contain the pivot table
            Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");

            // Create the pivot table on the pivot sheet
            int pivotIndex = pivotSheet.PivotTables.Add("A1:B5", "D1", "SalesPivot");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

            // Configure pivot fields: rows and data
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Category column
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1); // Amount column

            // Apply custom globalization settings that change the "Total" label
            workbook.Settings.GlobalizationSettings = new GlobalizationSettings();
            workbook.Settings.GlobalizationSettings.PivotSettings = new CustomPivotGlobalizationSettings();

            // Refresh and calculate the pivot table so the custom label takes effect
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("CustomTotalLabelDemo.xlsx");
        }

        // Custom subclass of PivotGlobalizationSettings to override the Total label text
        private class CustomPivotGlobalizationSettings : PivotGlobalizationSettings
        {
            // Return the desired custom text for the "Total" label
            public override string GetTextOfTotal()
            {
                return "My Custom Total";
            }
        }
    }
}
