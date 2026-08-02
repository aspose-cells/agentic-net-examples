// Title: Create a Pivot Table from a Dynamic Data Range with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to generate a new workbook, populate a source sheet, use MaxDisplayRange to capture the full data area, add a second worksheet, insert a pivot table that references the dynamic range, assign "Category" to rows and "Value" to data, refresh and calculate the pivot, and save the file as PivotTableDemo.xlsx.
// Keywords: Aspose.Cells pivot table C# | dynamic source range Aspose | MaxDisplayRange pivot | add pivot table .NET | Aspose.Cells example | Excel automation C# | programmatic pivot table
// Common Searches: how to add a pivot table with Aspose.Cells C# | use MaxDisplayRange as pivot source in Aspose.Cells | Aspose.Cells create pivot table from range | C# code for inserting pivot table in new worksheet | Aspose.Cells refresh and calculate pivot table
// Developer Intent: Programmatically create and save an Excel workbook that contains a pivot table built from a dynamically determined data range.
// Use Cases: Generate summary reports that automatically adapt to changing data sizes. | Automate Excel dashboards where row and data fields are defined in code. | Build reusable utilities that add pivot tables to any worksheet without hard‑coding ranges.
// AI Prompts: Write a reusable method that takes a worksheet name and adds a pivot table using its MaxDisplayRange as the source. | Show how to add multiple data fields to the pivot table and format the results as currency. | Explain the steps to refresh and recalculate a pivot table after updating the source data with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsPivotDemo
{
    // Demonstrates how to generate a new workbook, populate a source sheet, use MaxDisplayRange to capture the full data area, add a second worksheet, insert a pivot table that references the dynamic range, assign "Category" to rows and "Value" to data, refresh and calculate the pivot, and save the file as PivotTableDemo.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet (source data)
                Worksheet sourceSheet = workbook.Worksheets[0];
                sourceSheet.Name = "SourceData";

                // Populate sample data
                sourceSheet.Cells["A1"].PutValue("Category");
                sourceSheet.Cells["B1"].PutValue("Value");
                sourceSheet.Cells["A2"].PutValue("A");
                sourceSheet.Cells["B2"].PutValue(10);
                sourceSheet.Cells["A3"].PutValue("B");
                sourceSheet.Cells["B3"].PutValue(20);
                sourceSheet.Cells["A4"].PutValue("A");
                sourceSheet.Cells["B4"].PutValue(30);

                // Add a new worksheet to host the pivot table
                Worksheet pivotSheet = workbook.Worksheets.Add("PivotTable");

                // Determine the data range for the pivot table using MaxDisplayRange
                AsposeRange dataRange = sourceSheet.Cells.MaxDisplayRange;
                string sourceData = $"=SourceData!{dataRange.Address}";

                // Add the pivot table to the pivot sheet
                PivotTableCollection pivotTables = pivotSheet.PivotTables;
                int pivotIndex = pivotTables.Add(sourceData, "A1", "MyPivotTable");

                // Configure the pivot table (add fields)
                PivotTable pivotTable = pivotTables[pivotIndex];
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Value");

                // Refresh and calculate the pivot data
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Save the workbook
                workbook.Save("PivotTableDemo.xlsx");
                Console.WriteLine("Pivot table created and saved as PivotTableDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
