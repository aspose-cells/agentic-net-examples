// Title: C# – Customize Pivot Table Total Label with a Custom PivotGlobalizationSettings in Aspose.Cells
// Description: This example shows how to subclass SettablePivotGlobalizationSettings to replace the default "Total" label, assign the custom settings to a workbook, refresh and calculate the pivot table, and save the result as an Excel file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | PivotTable | Custom Total label | SettablePivotGlobalizationSettings | PivotGlobalizationSettings | GlobalizationSettings | RefreshData | CalculateData | Excel automation | Localization
// Common Searches: change total label Aspose.Cells pivot table | customize pivot table text C# Aspose | SettablePivotGlobalizationSettings example | apply custom globalization settings workbook | refresh pivot after globalization Aspose
// Developer Intent: Replace the default "Total" caption in an Aspose.Cells pivot table with a custom string by creating and applying a custom PivotGlobalizationSettings subclass.
// Use Cases: Brand a sales report with a company‑specific total caption instead of the generic "Total". | Localize Excel workbooks by providing language‑specific total labels through custom globalization classes. | Create multiple pivot tables in the same workbook, each showing a distinct total label for clearer data grouping.
// AI Prompts: Generate C# code that defines a SettablePivotGlobalizationSettings subclass to set the total label to "Grand Total" and applies it to an Aspose.Cells pivot table. | Explain how to refresh and recalculate a pivot table after assigning custom globalization settings in Aspose.Cells. | Show an example of localizing the total label of a pivot table using a custom PivotGlobalizationSettings class in a dynamic workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Settings;

namespace AsposeCellsExamples
{
    // Custom globalization settings that changes the "Total" label text
    // This example shows how to subclass SettablePivotGlobalizationSettings to replace the default "Total" label, assign the custom settings to a workbook, refresh and calculate the pivot table, and save the result as an Excel file using Aspose.Cells for .NET.
    public class CustomPivotGlobalizationSettings : SettablePivotGlobalizationSettings
    {
        public CustomPivotGlobalizationSettings()
        {
            // Set the desired text for the Total label
            SetTextOfTotal("Custom Total");
        }

        // Optionally override GetTextOfTotal to ensure the custom text is returned
        public override string GetTextOfTotal()
        {
            return base.GetTextOfTotal();
        }
    }

    public class ApplyCustomPivotGlobalization
    {
        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet dataSheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            dataSheet.Cells["A1"].PutValue("Category");
            dataSheet.Cells["B1"].PutValue("Amount");
            dataSheet.Cells["A2"].PutValue("Fruit");
            dataSheet.Cells["B2"].PutValue(120);
            dataSheet.Cells["A3"].PutValue("Vegetable");
            dataSheet.Cells["B3"].PutValue(80);
            dataSheet.Cells["A4"].PutValue("Fruit");
            dataSheet.Cells["B4"].PutValue(150);
            dataSheet.Cells["A5"].PutValue("Vegetable");
            dataSheet.Cells["B5"].PutValue(200);

            // Add a new worksheet for the pivot table
            Worksheet pivotSheet = workbook.Worksheets.Add("PivotTable");

            // Create the pivot table
            int pivotIndex = pivotSheet.PivotTables.Add("A1:B5", "D1", "SalesPivot");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

            // Configure pivot fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Category
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1); // Amount

            // Apply custom globalization settings
            GlobalizationSettings globalSettings = new GlobalizationSettings();
            globalSettings.PivotSettings = new CustomPivotGlobalizationSettings();
            workbook.Settings.GlobalizationSettings = globalSettings;

            // Refresh and calculate the pivot table to apply the custom label
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("CustomTotalLabelPivot.xlsx");
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
