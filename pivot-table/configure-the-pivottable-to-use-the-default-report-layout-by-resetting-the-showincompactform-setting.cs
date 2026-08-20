// Title: C# – Reset PivotTable to Default Compact Layout Using Aspose.Cells
// Description: Demonstrates how to create a workbook, add sample sales data, build a PivotTable, and call ShowInCompactForm() to reset the report layout to the default compact form. The example also refreshes, calculates, and saves the workbook as an .xlsx file.
// Keywords: Aspose.Cells | C# PivotTable | ShowInCompactForm | reset pivot layout | default compact form | Excel PivotTable programmatic | .NET Excel library | pivot table layout reset
// Common Searches: Aspose.Cells ShowInCompactForm example | reset pivot table layout C# | default compact form pivot table Aspose | how to set pivot table compact layout .NET | programmatically change pivot table layout Aspose.Cells
// Developer Intent: Reset a PivotTable’s report layout to the default compact form using Aspose.Cells for .NET.
// Use Cases: Generate an Excel report where the PivotTable always uses the standard compact layout for better readability. | Update source data, then refresh and recalculate the PivotTable while preserving its default layout. | Create a reusable C# routine that builds a PivotTable and ensures it starts in compact form before saving.
// AI Prompts: Write C# code with Aspose.Cells that creates a PivotTable and applies ShowInCompactForm to set the default layout. | Explain the impact of ShowInCompactForm on PivotTable appearance and how it interacts with RefreshData and CalculateData. | Provide a step‑by‑step guide to reset an existing PivotTable’s layout to compact form in an Aspose.Cells workbook.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, add sample sales data, build a PivotTable, and call ShowInCompactForm() to reset the report layout to the default compact form. The example also refreshes, calculates, and saves the workbook as an .xlsx file.
    public class ResetPivotTableLayout
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet (source data)
                Worksheet sourceSheet = workbook.Worksheets[0];
                sourceSheet.Name = "Data";

                // Populate sample data for the pivot table
                sourceSheet.Cells["A1"].PutValue("Category");
                sourceSheet.Cells["B1"].PutValue("Product");
                sourceSheet.Cells["C1"].PutValue("Sales");

                sourceSheet.Cells["A2"].PutValue("Electronics");
                sourceSheet.Cells["B2"].PutValue("Laptop");
                sourceSheet.Cells["C2"].PutValue(1200);

                sourceSheet.Cells["A3"].PutValue("Electronics");
                sourceSheet.Cells["B3"].PutValue("Phone");
                sourceSheet.Cells["C3"].PutValue(800);

                sourceSheet.Cells["A4"].PutValue("Furniture");
                sourceSheet.Cells["B4"].PutValue("Chair");
                sourceSheet.Cells["C4"].PutValue(150);

                sourceSheet.Cells["A5"].PutValue("Furniture");
                sourceSheet.Cells["B5"].PutValue("Table");
                sourceSheet.Cells["C5"].PutValue(300);

                // Add a new worksheet to host the pivot table
                Worksheet pivotSheet = workbook.Worksheets.Add("PivotTable");

                // Create the pivot table based on the source range
                int pivotIndex = pivotSheet.PivotTables.Add("=Data!A1:C5", "A3", "PivotTable1");
                PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

                // Add fields to the pivot table areas
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Reset the layout to the default (compact form)
                pivotTable.ShowInCompactForm();

                // Refresh pivot table data after source changes
                pivotTable.RefreshData();

                // Calculate the pivot table data
                pivotTable.CalculateData();

                // Determine output file path
                string outputFile = "PivotTableDefaultLayout.xlsx";
                string outputPath = Path.GetFullPath(outputFile);

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
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
            ResetPivotTableLayout.Run();
        }
    }
}
