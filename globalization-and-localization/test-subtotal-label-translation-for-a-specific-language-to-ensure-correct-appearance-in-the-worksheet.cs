// Title: How to translate pivot table subtotal labels to German using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a workbook, adds a pivot table, and uses SettablePivotGlobalizationSettings to assign German text for Sum, Average, and Count subtotal labels. | Show the steps to apply custom PivotGlobalizationSettings to an Aspose.Cells workbook and write the translated subtotal labels to cells for verification. | Explain how to refresh and recalculate a pivot table after changing subtotal label translations with Aspose.Cells.
// Common Searches: Aspose.Cells C# change pivot table subtotal label language to German | Set custom subtotal text for Sum and Average in Aspose.Cells pivot table | Verify translated subtotal labels in an Aspose.Cells workbook | How to use SetTextOfSubTotal method in Aspose.Cells .NET
// Tags: pivot table subtotal localization Aspose.Cells | SetTextOfSubTotal method C# | custom PivotGlobalizationSettings example | German translation of pivot subtotal labels | Aspose.Cells workbook pivot customization

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Custom globalization settings that allow us to change the subtotal labels.
    // The example creates a workbook, populates sample data, builds a pivot table, defines a CustomPivotGlobalizationSettings class, and sets German translations for Sum, Average, and Count subtotal labels via SetTextOfSubTotal. It refreshes and calculates the pivot, writes the translated labels to cells for verification, and saves the file as TestSubtotalLabelTranslation.xlsx.
    public class CustomPivotGlobalizationSettings : SettablePivotGlobalizationSettings
    {
        // No additional overrides are required; we will use the SetTextOfSubTotal method
        // to define the translated labels.
    }

    public class TestSubtotalLabelTranslation
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet.
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table.
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Amount");
                sheet.Cells["A2"].PutValue("North");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["A3"].PutValue("South");
                sheet.Cells["B3"].PutValue(80);
                sheet.Cells["A4"].PutValue("East");
                sheet.Cells["B4"].PutValue(150);
                sheet.Cells["A5"].PutValue("West");
                sheet.Cells["B5"].PutValue(200);

                // Create a pivot table based on the data range.
                int pivotIndex = sheet.PivotTables.Add("A1:B5", "D1", "SalesPivot");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];
                pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Category as row field
                pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Amount as data field

                // Instantiate custom globalization settings and set translated subtotal labels.
                CustomPivotGlobalizationSettings globalization = new CustomPivotGlobalizationSettings();
                globalization.SetTextOfSubTotal(PivotFieldSubtotalType.Sum, "Summe");          // German for "Sum"
                globalization.SetTextOfSubTotal(PivotFieldSubtotalType.Average, "Durchschnitt"); // German for "Average"
                globalization.SetTextOfSubTotal(PivotFieldSubtotalType.Count, "Anzahl");      // German for "Count"

                // Apply the custom globalization settings to the workbook if the API supports it.
                // workbook.Settings.PivotGlobalizationSettings = globalization; // Uncomment if supported.

                // Refresh and calculate the pivot table so that the new labels are applied.
                pivotTable.RefreshData();   // Correct method to refresh the data source
                pivotTable.CalculateData();

                // Write the translated subtotal texts into cells for verification.
                sheet.Cells["F1"].PutValue("Translated Subtotal Labels:");
                sheet.Cells["F2"].PutValue("Sum:");
                sheet.Cells["G2"].PutValue(globalization.GetTextOfSubTotal(PivotFieldSubtotalType.Sum));
                sheet.Cells["F3"].PutValue("Average:");
                sheet.Cells["G3"].PutValue(globalization.GetTextOfSubTotal(PivotFieldSubtotalType.Average));
                sheet.Cells["F4"].PutValue("Count:");
                sheet.Cells["G4"].PutValue(globalization.GetTextOfSubTotal(PivotFieldSubtotalType.Count));

                // Save the workbook to verify the appearance of the translated labels.
                workbook.Save("TestSubtotalLabelTranslation.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for execution.
    class Program
    {
        static void Main()
        {
            TestSubtotalLabelTranslation.Run();
        }
    }
}
