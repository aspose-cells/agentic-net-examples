// Title: Set ConsolidationFunction.Sum for a data field in an Aspose.Cells PivotTable (C# example)
// AI Prompts: Generate C# code that creates a PivotTable from a range, adds a data field, and sets its ConsolidationFunction to Sum using Aspose.Cells. | Show how to refresh and calculate an Aspose.Cells PivotTable after changing a data field's aggregation to Sum. | Provide a step‑by‑step example of adding row and data fields to a PivotTable and configuring the data field to use the Sum function in C#.
// Common Searches: Aspose.Cells C# set pivot table data field aggregation to sum | How to use ConsolidationFunction.Sum with Aspose.Cells PivotTable | Refresh and calculate pivot table after setting function in Aspose.Cells .NET | Add row field and sum data field to pivot table using Aspose.Cells API | C# example for configuring sum consolidation in Aspose.Cells pivot cache
// Tags: Aspose.Cells PivotTable set data field sum | C# ConsolidationFunction.Sum usage | refresh calculate pivot table Aspose.Cells | add row and data fields Aspose.Cells | pivot table aggregation function .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Creates a workbook, fills it with product and sales data, adds a PivotTable on range A1:B5, assigns 'Product' as a row field, adds 'Sales' as a data field, sets its ConsolidationFunction to Sum, refreshes and calculates the pivot, and saves the result as PivotFieldSetSumFunctionDemo.xlsx.
    public class PivotFieldSetSumFunctionDemo
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
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

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

            // Add a pivot table based on the data range
            PivotTableCollection pivotTables = worksheet.PivotTables;
            int pivotIndex = pivotTables.Add("A1:B5", "E3", "PivotTable1");
            PivotTable pivotTable = pivotTables[pivotIndex];

            // Add a row field (Product)
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");

            // Add a data field (Sales) and set its consolidation function to Sum
            int dataFieldIndex = pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");
            PivotField dataField = pivotTable.DataFields[dataFieldIndex];
            dataField.Function = ConsolidationFunction.Sum; // Set aggregation to Sum

            // Refresh and calculate the pivot table data using the correct API
            pivotTable.RefreshData();   // Refreshes the pivot cache
            pivotTable.CalculateData(); // Calculates the pivot table values

            // Save the workbook
            workbook.Save("PivotFieldSetSumFunctionDemo.xlsx");
        }
    }
}
