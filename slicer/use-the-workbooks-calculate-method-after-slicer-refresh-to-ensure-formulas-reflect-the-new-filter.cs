// Title: Refresh a Pivot Table Slicer and Recalculate Workbook Formulas with Aspose.Cells in C#
// AI Prompts: Refresh the slicer linked to a pivot table, then invoke workbook.CalculateFormula to update all dependent formulas in an Aspose.Cells workbook using C#. | After modifying source data, call slicer.Refresh followed by workbook.CalculateFormula so that pivot totals and other calculations reflect the new values.
// Common Searches: Aspose.Cells C# recalculate formulas after slicer refresh | how to update pivot totals when slicer changes using Aspose.Cells | using workbook.CalculateFormula with slicer.Refresh in .NET | example code for slicer.Refresh and formula recalculation Aspose.Cells | C# Aspose.Cells refresh slicer linked to pivot table
// Tags: slicer linked pivot refresh Aspose.Cells C# | invoke workbook.CalculateFormula after slicer update | update pivot totals via slicer Aspose.Cells | recalculate all workbook formulas C# Aspose.Cells | Aspose.Cells slicer example with pivot table

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

// Demonstrates creating a workbook with sample data, adding a pivot table and a linked slicer, changing source values, refreshing the slicer (which also refreshes the pivot), recalculating all formulas with CalculateFormula, and saving the updated file.
class SlicerRefreshCalculateDemo
{
    static void Main(string[] args)
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
        // Create a new workbook and add sample data
        Workbook workbook = new Workbook();
        Worksheet dataSheet = workbook.Worksheets[0];
        dataSheet.Cells["A1"].PutValue("Product");
        dataSheet.Cells["B1"].PutValue("Sales");
        dataSheet.Cells["A2"].PutValue("Apple");
        dataSheet.Cells["B2"].PutValue(100);
        dataSheet.Cells["A3"].PutValue("Banana");
        dataSheet.Cells["B3"].PutValue(200);

        // Add a formula that sums the sales column
        dataSheet.Cells["C1"].PutValue("TotalSales");
        dataSheet.Cells["C2"].Formula = "=SUM(B2:B3)";

        // Create a pivot table based on the data range
        Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");
        int pivotIndex = pivotSheet.PivotTables.Add("A1:B3", "D3", "PivotTable1");
        PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];
        pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Product
        pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Sales

        // Add a slicer linked to the pivot table
        Worksheet slicerSheet = workbook.Worksheets.Add("Slicer");
        int slicerIndex = slicerSheet.Slicers.Add(pivotTable, "A1", "Product");
        Slicer slicer = slicerSheet.Slicers[slicerIndex];

        // Modify source data to demonstrate slicer refresh effect
        dataSheet.Cells["A2"].PutValue("Orange");
        dataSheet.Cells["B2"].PutValue(150);

        // Refresh the slicer (this also refreshes the pivot table)
        slicer.Refresh();

        // Recalculate all formulas in the workbook so that the total reflects the refreshed data
        workbook.CalculateFormula();

        // Save the workbook
        workbook.Save("SlicerRefreshCalculateDemo.xlsx");
    }
}
