// Title: C# – Insert a Pivot Table with a Defined Source Range using Aspose.Cells for .NET
// Description: Creates a new workbook, adds sample data to a "SourceData" sheet, builds a range string that includes the sheet name, inserts a pivot table on a separate "PivotTable" sheet, assigns "Category" to rows and "Value" to data, refreshes the cache, calculates results, and saves the file as PivotTableDemo.xlsx.
// Keywords: Aspose.Cells | Aspose.Cells for .NET | C# pivot table | pivot table example | define source range | add pivot table programmatically | PivotTableCollection | PivotTable API | GitHub sample | Aspose.Cells tutorial
// Common Searches: Aspose.Cells create pivot table from range C# | How to add pivot table to new worksheet using Aspose.Cells .NET | Refresh pivot table data Aspose.Cells C# | GitHub Aspose.Cells pivot table example | Dynamic source range pivot table Aspose.Cells
// Developer Intent: Programmatically add and configure a pivot table in a .NET workbook using a specific source data range.
// Use Cases: Generate a sales summary that groups categories and totals values automatically. | Automate financial consolidation by pivoting raw transaction data into aggregated rows and columns. | Build a dynamic dashboard workbook where the pivot table updates whenever the source data changes.
// AI Prompts: Write C# code with Aspose.Cells to insert a pivot table that includes multiple data fields and a filter field. | Show how to change the source range of an existing pivot table and refresh its data using Aspose.Cells. | Provide an example of adding calculated fields to a pivot table created with Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Creates a new workbook, adds sample data to a "SourceData" sheet, builds a range string that includes the sheet name, inserts a pivot table on a separate "PivotTable" sheet, assigns "Category" to rows and "Value" to data, refreshes the cache, calculates results, and saves the file as PivotTableDemo.xlsx.
class InsertPivotTable
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Prepare source data worksheet
            Worksheet sourceSheet = workbook.Worksheets[0];
            sourceSheet.Name = "SourceData";

            // Add sample data
            sourceSheet.Cells["A1"].PutValue("Category");
            sourceSheet.Cells["B1"].PutValue("Value");
            sourceSheet.Cells["A2"].PutValue("A");
            sourceSheet.Cells["B2"].PutValue(10);
            sourceSheet.Cells["A3"].PutValue("B");
            sourceSheet.Cells["B3"].PutValue(20);
            sourceSheet.Cells["A4"].PutValue("A");
            sourceSheet.Cells["B4"].PutValue(30);

            // Add a worksheet that will contain the pivot table
            Worksheet pivotSheet = workbook.Worksheets.Add("PivotTable");

            // Build the source data range string (including sheet name)
            var srcRange = sourceSheet.Cells.MaxDisplayRange;
            string sourceData = $"{sourceSheet.Name}!{srcRange.Address}";

            // Insert the pivot table
            PivotTableCollection pivots = pivotSheet.PivotTables;
            int pivotIndex = pivots.Add(sourceData, "A1", "MyPivotTable");

            // Configure the pivot table fields
            PivotTable pivot = pivots[pivotIndex];
            pivot.AddFieldToArea(PivotFieldType.Row, "Category");
            pivot.AddFieldToArea(PivotFieldType.Data, "Value");

            // Refresh the pivot cache and calculate data
            pivot.RefreshData();   // refreshes source data
            pivot.CalculateData(); // calculates the pivot values

            // Save the workbook
            workbook.Save("PivotTableDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
