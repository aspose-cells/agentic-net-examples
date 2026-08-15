// Title: Aspose.Cells C# Example: Set PivotTable Data Field to Average (ConsolidationFunction)
// Description: This C# sample creates a workbook, adds sample data, builds a PivotTable, places the "Category" column in the row area, adds the "Amount" column as a data field, changes its ConsolidationFunction to Average, refreshes and recalculates the pivot, and saves the result as PivotTable_AverageFunction.xlsx.
// Keywords: Aspose.Cells | C# | .NET | PivotTable | Average | ConsolidationFunction | data field function | refresh pivot | calculate pivot | Excel automation | GitHub example | sample code
// Common Searches: Aspose.Cells set pivot table average C# | Change consolidation function to Average in Aspose.Cells | PivotTable average calculation .NET | How to use ConsolidationFunction.Average with Aspose.Cells | C# example for averaging pivot table values
// Developer Intent: Apply the Average consolidation function to a PivotTable data field using Aspose.Cells for .NET.
// Use Cases: Generate a report showing average sales per category directly from raw data. | Create financial dashboards where expense averages are displayed instead of totals. | Automate Excel analytics that require mean performance metrics calculated via PivotTables.
// AI Prompts: Write C# code that creates a PivotTable with Aspose.Cells and sets the data field's ConsolidationFunction to Average. | Explain how to refresh and recalculate a PivotTable after changing its consolidation function in Aspose.Cells for .NET. | Provide steps to save a workbook after applying an Average function to a PivotTable data field using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // This C# sample creates a workbook, adds sample data, builds a PivotTable, places the "Category" column in the row area, adds the "Amount" column as a data field, changes its ConsolidationFunction to Average, refreshes and recalculates the pivot, and saves the result as PivotTable_AverageFunction.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Populate sample data for the pivot table
                cells["A1"].Value = "Category";
                cells["B1"].Value = "Amount";
                cells["A2"].Value = "A";
                cells["B2"].Value = 100;
                cells["A3"].Value = "B";
                cells["B3"].Value = 150;
                cells["A4"].Value = "A";
                cells["B4"].Value = 200;
                cells["A5"].Value = "B";
                cells["B5"].Value = 250;

                // Add a pivot table based on the data range
                PivotTableCollection pivotTables = worksheet.PivotTables;
                int pivotIndex = pivotTables.Add("A1:B5", "D3", "PivotTable1");
                PivotTable pivotTable = pivotTables[pivotIndex];

                // Add fields to the pivot table
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");

                // Add the data field and set its consolidation function to Average
                int dataFieldIndex = pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");
                PivotField dataField = pivotTable.DataFields[dataFieldIndex];
                dataField.Function = ConsolidationFunction.Average;

                // Refresh and calculate the pivot table to apply the new function
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Save the workbook
                string outputPath = "PivotTable_AverageFunction.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
