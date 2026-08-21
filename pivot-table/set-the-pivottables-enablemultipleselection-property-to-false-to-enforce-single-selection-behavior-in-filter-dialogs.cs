// Title: Aspose.Cells .NET – Disable multi‑select in PivotTable filter dialogs with EnableMultipleSelection = false
// Description: Shows how to create a workbook, populate it with sample sales data, add a PivotTable, and set the PivotTable.EnableMultipleSelection (or AllowMultipleFiltersPerField) property to false so that filter dialogs permit only one item. The code refreshes the cache, calculates the data, and saves the workbook as an XLSX file.
// Keywords: Aspose.Cells | PivotTable | EnableMultipleSelection | AllowMultipleFiltersPerField | single selection filter | disable multi‑select | .NET | C# | Excel pivot table | filter dialog | Aspose.Cells example
// Common Searches: Aspose.Cells set EnableMultipleSelection false | disable multi select in PivotTable filter Aspose.Cells | AllowMultipleFiltersPerField false example .NET | single selection filter dialog Aspose.Cells PivotTable | how to restrict PivotTable filters to one choice using Aspose.Cells
// Developer Intent: Turn off multi‑select in PivotTable filter dialogs to enforce a single‑selection UI.
// Use Cases: Generate a sales report where users can pick only one category at a time in the PivotTable filter. | Create an interactive dashboard that limits each filter field to a single selection, simplifying data analysis. | Programmatically configure a PivotTable before exporting to ensure compliance with UI guidelines that prohibit multiple selections.
// AI Prompts: Provide a C# example that creates a PivotTable with Aspose.Cells and disables multi‑select in its filter dialogs using EnableMultipleSelection = false. | Explain the impact of setting AllowMultipleFiltersPerField to false on PivotTable filter behavior in Aspose.Cells. | Show how to refresh and calculate a PivotTable after changing its filter selection mode with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Shows how to create a workbook, populate it with sample sales data, add a PivotTable, and set the PivotTable.EnableMultipleSelection (or AllowMultipleFiltersPerField) property to false so that filter dialogs permit only one item. The code refreshes the cache, calculates the data, and saves the workbook as an XLSX file.
    public class PivotTableEnableMultipleSelectionDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                sheet.Cells["A1"].Value = "Product";
                sheet.Cells["B1"].Value = "Category";
                sheet.Cells["C1"].Value = "Sales";

                sheet.Cells["A2"].Value = "Apple";
                sheet.Cells["B2"].Value = "Fruit";
                sheet.Cells["C2"].Value = 1000;

                sheet.Cells["A3"].Value = "Banana";
                sheet.Cells["B3"].Value = "Fruit";
                sheet.Cells["C3"].Value = 1500;

                sheet.Cells["A4"].Value = "Carrot";
                sheet.Cells["B4"].Value = "Vegetable";
                sheet.Cells["C4"].Value = 800;

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:C4", "E3", "SalesPivot");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Add fields to the pivot table
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
                pivotTable.AddFieldToArea(PivotFieldType.Column, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Enforce single selection behavior in filter dialogs
                pivotTable.AllowMultipleFiltersPerField = false;

                // Refresh and calculate the pivot table to apply changes
                pivotTable.RefreshData();   // Correct API to refresh cache
                pivotTable.CalculateData();

                // Save the workbook
                string outputPath = "PivotTable_EnableMultipleSelection_Demo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            PivotTableEnableMultipleSelectionDemo.Run();
        }
    }
}
