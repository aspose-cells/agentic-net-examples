// Title: Aspose.Cells C# – Set PivotTable Data Field ConsolidationFunction to Sum
// Description: C# code that builds a workbook with product‑sales rows, creates a PivotTable on A1:B5, adds "Product" as a row field, assigns "Sales" as a data field, changes its ConsolidationFunction to Sum, refreshes the cache, calculates the results, and saves the file as PivotFieldFunctionSum.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | PivotTable | ConsolidationFunction.Sum | set pivot field function | sum aggregation | Excel pivot programmatically | pivot cache refresh | data field function
// Common Searches: How to set ConsolidationFunction.Sum for a PivotTable in Aspose.Cells C# | Aspose.Cells example for summing pivot table values | Change pivot data field aggregation to Sum using .NET | C# code to configure PivotTable sum function with Aspose.Cells | Programmatic Excel pivot table sum aggregation Aspose
// Developer Intent: Configure a PivotTable’s data field to use the Sum aggregation so numeric entries are totaled correctly.
// Use Cases: Generate sales reports that total revenue per product automatically. | Create financial dashboards where expense categories are summed across periods. | Build Excel workbooks that summarize large data sets with a single‑click sum pivot. | Automate data‑analysis pipelines that require summed metrics in pivot tables.
// AI Prompts: Write C# code with Aspose.Cells to create a PivotTable and set its data field function to Sum. | Explain the effect of ConsolidationFunction.Sum on PivotTable calculations in Aspose.Cells. | Show how to modify an existing Aspose.Cells PivotTable to change the aggregation method to Sum.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // C# code that builds a workbook with product‑sales rows, creates a PivotTable on A1:B5, adds "Product" as a row field, assigns "Sales" as a data field, changes its ConsolidationFunction to Sum, refreshes the cache, calculates the results, and saves the file as PivotFieldFunctionSum.xlsx using Aspose.Cells for .NET.
    class SetPivotFieldFunctionSum
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data for the pivot table
                cells["A1"].Value = "Product";
                cells["B1"].Value = "Sales";
                cells["A2"].Value = "A";
                cells["B2"].Value = 100;
                cells["A3"].Value = "B";
                cells["B3"].Value = 120;
                cells["A4"].Value = "A";
                cells["B4"].Value = 80;
                cells["A5"].Value = "B";
                cells["B5"].Value = 60;

                // Add a pivot table using the data range A1:B5, place it at E3, and name it "PivotTable1"
                int pivotIndex = sheet.PivotTables.Add("A1:B5", "E3", "PivotTable1");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Add "Product" as a row field
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");

                // Add "Sales" as a data field and retrieve the created PivotField object
                int dataFieldIndex = pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");
                PivotField dataField = pivotTable.DataFields[dataFieldIndex];

                // Set the consolidation function of the data field to Sum
                dataField.Function = ConsolidationFunction.Sum;

                // Refresh pivot cache data and calculate the results
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Save the workbook with the configured pivot table
                workbook.Save("PivotFieldFunctionSum.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            SetPivotFieldFunctionSum.Run();
        }
    }
}
