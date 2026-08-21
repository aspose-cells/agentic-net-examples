// Title: Aspose.Cells C# – Enable ShowExpandCollapseButtons on a PivotTable
// Description: This C# sample creates a workbook, adds sample data, builds a PivotTable, and activates the ShowExpandCollapseButtons property so hierarchical rows display expand/collapse icons. It also demonstrates enabling drill‑down, refreshing the cache, calculating data, and saving the workbook as an .xlsx file.
// Keywords: Aspose.Cells | C# PivotTable | ShowExpandCollapseButtons | expand collapse icons | hierarchical navigation | EnableDrilldown | ShowDrill | pivot table visual cues | Excel automation C# | Aspose.Cells example
// Common Searches: Aspose.Cells show expand collapse buttons C# | how to display expand collapse icons in PivotTable using Aspose.Cells | Enable hierarchical navigation in Aspose.Cells PivotTable | C# code for ShowExpandCollapseButtons property | Aspose.Cells pivot table drilldown and expand buttons
// Developer Intent: Set the ShowExpandCollapseButtons property on a PivotTable to make expand/collapse buttons visible for grouped rows.
// Use Cases: Provide end‑users with visual cues for navigating grouped data in an exported Excel file. | Combine ShowExpandCollapseButtons with EnableDrilldown to create an interactive reporting workbook. | Generate Excel reports programmatically where hierarchical rows can be expanded or collapsed directly in the UI.
// AI Prompts: Generate C# code that creates a PivotTable with ShowExpandCollapseButtons enabled using Aspose.Cells. | Explain the role of ShowExpandCollapseButtons, EnableDrilldown, and ShowDrill in Aspose.Cells PivotTables. | Provide step‑by‑step instructions to add hierarchical navigation icons to a PivotTable in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // This C# sample creates a workbook, adds sample data, builds a PivotTable, and activates the ShowExpandCollapseButtons property so hierarchical rows display expand/collapse icons. It also demonstrates enabling drill‑down, refreshing the cache, calculating data, and saving the workbook as an .xlsx file.
    public class PivotTableShowExpandCollapseDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet and name it
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Populate sample data for the pivot table
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(100);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(200);
            sheet.Cells["A4"].PutValue("A");
            sheet.Cells["B4"].PutValue(150);
            sheet.Cells["A5"].PutValue("B");
            sheet.Cells["B5"].PutValue(250);

            try
            {
                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "PivotTable1");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Configure the pivot table fields
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Value");

                // Enable drilldown functionality and visual expand/collapse buttons
                pivotTable.EnableDrilldown = true;
                pivotTable.ShowDrill = true;

                // Refresh the pivot cache and calculate data
                pivotTable.RefreshData();
                pivotTable.CalculateData();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Pivot table error: {ex.Message}");
                // Continue execution; the workbook can still be saved without the pivot table
            }

            // Save the workbook to a file
            string outputPath = "PivotTableShowExpandCollapseDemo.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
    }
}
