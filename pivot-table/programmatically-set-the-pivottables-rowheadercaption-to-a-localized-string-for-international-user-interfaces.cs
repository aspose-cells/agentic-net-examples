// Title: How to set a localized RowHeaderCaption and row label text for an Aspose.Cells PivotTable in C#
// AI Prompts: Generate C# code using Aspose.Cells that creates a workbook, adds a pivot table, and assigns a French string to the PivotTable.RowHeaderCaption property. | Show how to apply SettablePivotGlobalizationSettings in Aspose.Cells to replace the default "Row Labels" caption with a custom localized label before saving the workbook.
// Common Searches: aspnet c# set pivot table row header caption to French using Aspose.Cells | localize "Row Labels" text in Aspose.Cells pivot tables with SettablePivotGlobalizationSettings | Aspose.Cells example for customizing pivot table row header caption for international users | how to change pivot table row header caption programmatically in C# Aspose.Cells | apply globalization settings to pivot tables in Aspose.Cells workbook
// Tags: Aspose.Cells pivot table header text localization | SettablePivotGlobalizationSettings row labels translation | C# customize Aspose.Cells pivot table captions | Excel workbook internationalization with Aspose.Cells | Aspose.Cells save workbook with localized pivot labels

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Settings;

// The example creates a workbook with sample data, adds a pivot table, sets the RowHeaderCaption to a French label, uses SettablePivotGlobalizationSettings to replace the default "Row Labels" text with a localized version, and saves the file as PivotRowHeaderLocalized.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and add a data worksheet
        Workbook workbook = new Workbook();
        Worksheet dataSheet = workbook.Worksheets[0];
        dataSheet.Name = "Data";

        // Populate sample data for the pivot table
        var cells = dataSheet.Cells;
        cells["A1"].PutValue("Country");
        cells["B1"].PutValue("Sales");
        cells["A2"].PutValue("France");
        cells["B2"].PutValue(150);
        cells["A3"].PutValue("Germany");
        cells["B3"].PutValue(200);

        // Add a separate worksheet that will contain the pivot table
        Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");

        // Create the pivot table (source range is the data worksheet)
        int pivotIndex = pivotSheet.PivotTables.Add("=Data!A1:B3", "A3", "SalesPivot");
        PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

        // Set the RowHeaderCaption to a localized string (e.g., French)
        pivotTable.RowHeaderCaption = "Pays"; // "Country" in French

        // Configure the pivot fields
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Country");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

        // OPTIONAL: Use SettablePivotGlobalizationSettings to localize the generic "Row Labels" text
        SettablePivotGlobalizationSettings pivotGlobalization = new SettablePivotGlobalizationSettings();
        pivotGlobalization.SetTextOfRowLabels("Étiquettes de ligne"); // localized label

        // Wrap the pivot globalization settings into a SettableGlobalizationSettings instance
        SettableGlobalizationSettings globalization = new SettableGlobalizationSettings();
        globalization.PivotSettings = pivotGlobalization;

        // Apply the globalization settings to the workbook
        workbook.Settings.GlobalizationSettings = globalization;

        // Save the workbook
        workbook.Save("PivotRowHeaderLocalized.xlsx", SaveFormat.Xlsx);
    }
}
