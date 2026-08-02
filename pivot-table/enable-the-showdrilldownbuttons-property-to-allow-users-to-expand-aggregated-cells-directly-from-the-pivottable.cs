// Title: C# – Enable ShowDrill (expand/collapse) Buttons on an Aspose.Cells PivotTable
// Description: Demonstrates how to create a workbook, add sample data, insert a PivotTable, assign row and data fields, and activate drill‑down icons by setting EnableDrilldown and ShowDrill to true. The example refreshes, calculates, and saves the file as PivotTableShowDrillDemo.xlsx.
// Keywords: Aspose.Cells PivotTable ShowDrill | EnableDrilldown C# | Aspose.Cells drill down buttons | expand collapse icons PivotTable | .NET Excel pivot table example | Aspose.Cells interactive PivotTable | C# Excel pivot drill‑down
// Common Searches: Aspose.Cells show drill buttons on pivot table | Enable drill‑down icons in PivotTable using .NET | C# code for ShowDrill property Aspose.Cells | How to add expand collapse buttons to Excel pivot with Aspose | PivotTable EnableDrilldown example C#
// Developer Intent: Add interactive expand/collapse icons to a PivotTable so end users can drill into aggregated values directly from the generated Excel file.
// Use Cases: Building automated reports that let users explore detailed rows by clicking drill‑down icons. | Creating Excel dashboards with interactive PivotTables for on‑the‑fly data analysis. | Generating workbooks where category totals need to be expandable without manual Excel configuration.
// AI Prompts: Write C# code using Aspose.Cells to create a PivotTable and turn on both EnableDrilldown and ShowDrill properties. | Explain the visual effect of the ShowDrill property in Excel and how to set it with Aspose.Cells. | Provide a step‑by‑step tutorial for adding drill‑down buttons to a PivotTable, refreshing data, and saving the workbook in .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, add sample data, insert a PivotTable, assign row and data fields, and activate drill‑down icons by setting EnableDrilldown and ShowDrill to true. The example refreshes, calculates, and saves the file as PivotTableShowDrillDemo.xlsx.
    public class PivotTableShowDrillDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet and give it a name
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Data";

                // Populate sample data for the pivot table
                sheet.Cells["A1"].Value = "Category";
                sheet.Cells["B1"].Value = "Value";
                sheet.Cells["A2"].Value = "A";
                sheet.Cells["B2"].Value = 100;
                sheet.Cells["A3"].Value = "B";
                sheet.Cells["B3"].Value = 200;
                sheet.Cells["A4"].Value = "A";
                sheet.Cells["B4"].Value = 150;

                // Add a pivot table to the worksheet
                int pivotIndex = sheet.PivotTables.Add("A1:B4", "E3", "PivotTable1");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Configure the pivot table fields
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Value");

                // Enable drilldown functionality and show the expand/collapse buttons
                pivotTable.EnableDrilldown = true;
                pivotTable.ShowDrill = true;

                // Refresh data and calculate the pivot table
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Save the workbook to a file
                string outputPath = "PivotTableShowDrillDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            PivotTableShowDrillDemo.Run();
        }
    }
}
