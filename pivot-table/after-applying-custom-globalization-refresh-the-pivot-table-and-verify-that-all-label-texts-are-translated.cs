// Title: C# – Apply Custom Globalization to an Aspose.Cells Pivot Table, Refresh It, and Verify Translated Labels
// Description: Creates a workbook with sample sales data, adds a pivot table, defines Spanish equivalents for row labels, column labels, total, grand total, (All) and data field header using SettablePivotGlobalizationSettings, attaches the settings to the workbook, refreshes and calculates the pivot table to apply the changes, reads back the custom texts for verification, and saves the file.
// Keywords: Aspose.Cells | C# | PivotTable | globalization | custom label translation | Spanish localization | SettablePivotGlobalizationSettings | RefreshData | CalculateData | Excel automation
// Common Searches: Aspose.Cells change pivot table labels to Spanish | Refresh pivot table after setting custom globalization C# | Verify custom pivot table texts in Aspose.Cells workbook | Set custom row and column label text for Aspose.Cells pivot table | Localize Aspose.Cells pivot table UI elements
// Developer Intent: Apply custom globalization strings to a pivot table, refresh the table, and confirm that the label texts are displayed in the target language.
// Use Cases: Generate a workbook with sample data and a pivot table, then replace default UI labels with Spanish equivalents. | Refresh and recalculate the pivot table after modifying GlobalizationSettings so the new labels appear in the worksheet. | Read back the applied texts via GetText methods to programmatically verify successful localization before saving.
// AI Prompts: Write C# code that sets Spanish globalization texts for an Aspose.Cells pivot table and validates the changes. | Show how to refresh and calculate a pivot table after updating GlobalizationSettings, then retrieve the applied label texts. | Explain step‑by‑step how SettablePivotGlobalizationSettings is used to localize pivot table UI elements in a .NET workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Settings;

namespace AsposeCellsPivotGlobalizationDemo
{
    // Creates a workbook with sample sales data, adds a pivot table, defines Spanish equivalents for row labels, column labels, total, grand total, (All) and data field header using SettablePivotGlobalizationSettings, attaches the settings to the workbook, refreshes and calculates the pivot table to apply the changes, reads back the custom texts for verification, and saves the file.
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

            // Configure the pivot fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);      // Product as row field
            pivotTable.AddFieldToArea(PivotFieldType.Column, 1);   // Region as column field
            pivotTable.AddFieldToArea(PivotFieldType.Data, 2);     // Sales as data field

            // Create customizable globalization settings for the pivot table
            SettablePivotGlobalizationSettings pivotSettings = new SettablePivotGlobalizationSettings();

            // Set custom texts for various pivot table labels
            pivotSettings.SetTextOfRowLabels("Filas");                     // Custom "Row Labels"
            pivotSettings.SetTextOfColumnLabels("Columnas");               // Custom "Column Labels"
            pivotSettings.SetTextOfTotal("Total Personalizado");          // Custom "Total"
            pivotSettings.SetTextOfGrandTotal("Gran Total");              // Custom "Grand Total"
            pivotSettings.SetTextOfAll("(Todo)");                         // Custom "(All)"
            pivotSettings.SetTextOfDataFieldHeader("Valores");            // Custom data field header

            // Attach the custom settings to the workbook
            GlobalizationSettings globalSettings = new GlobalizationSettings();
            globalSettings.PivotSettings = pivotSettings;
            workbook.Settings.GlobalizationSettings = globalSettings;

            // Refresh and calculate the pivot table to apply the custom globalization
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Verify that the custom texts are applied by reading them from the settings
            Console.WriteLine("Verification of applied globalization texts:");
            Console.WriteLine($"Row Labels text   : {pivotSettings.GetTextOfRowLabels()}");
            Console.WriteLine($"Column Labels text: {pivotSettings.GetTextOfColumnLabels()}");
            Console.WriteLine($"Total text        : {pivotSettings.GetTextOfTotal()}");
            Console.WriteLine($"Grand Total text  : {pivotSettings.GetTextOfGrandTotal()}");
            Console.WriteLine($"(All) text        : {pivotSettings.GetTextOfAll()}");
            Console.WriteLine($"Data Field Header : {pivotSettings.GetTextOfDataFieldHeader()}");

            // Save the workbook
            workbook.Save("PivotTableWithCustomGlobalization.xlsx");
        }
    }
}
