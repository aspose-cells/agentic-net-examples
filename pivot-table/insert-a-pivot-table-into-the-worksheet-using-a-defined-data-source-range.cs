// Title: Create a pivot table from a defined source range on a new worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that uses Aspose.Cells to define a source range, add a pivot table to a separate worksheet, and set row and data fields. | Show how to refresh the pivot cache and calculate data for a pivot table created with Aspose.Cells. | Provide the steps to save the workbook containing the pivot table as an .xlsx file with Aspose.Cells.
// Common Searches: Aspose.Cells C# how to add a pivot table from a specific cell range | example of setting row and data fields for a pivot table with Aspose.Cells | refreshing and calculating pivot table data in Aspose.Cells .NET | saving a workbook with a pivot table to Excel format using Aspose.Cells
// Tags: Aspose.Cells add pivot table from range | Aspose.Cells configure pivot row field | Aspose.Cells refresh pivot cache | Aspose.Cells save workbook as xlsx | Aspose.Cells pivot table data calculation

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using AsposeRange = Aspose.Cells.Range;

namespace PivotTableDemo
{
    // The sample creates a workbook, populates a source sheet with sample data, determines the full data range, adds a new worksheet, inserts a pivot table named "MyPivotTable" based on that range, assigns "Category" as a row field and "Value" as a data field, refreshes and calculates the pivot, and finally saves the workbook as PivotTableDemo.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet and add sample data
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

                // Add a new worksheet that will contain the pivot table
                Worksheet pivotSheet = workbook.Worksheets.Add("PivotTable");

                // Determine the source data range (including the header row)
                AsposeRange sourceRange = sourceSheet.Cells.MaxDisplayRange;
                string sourceData = $"=SourceData!{sourceRange.Address}";

                // Add the pivot table to the pivot sheet
                PivotTableCollection pivotTables = pivotSheet.PivotTables;
                int pivotIndex = pivotTables.Add(sourceData, "A1", "MyPivotTable");

                // Retrieve the created pivot table
                PivotTable pivotTable = pivotTables[pivotIndex];

                // Configure the pivot table fields
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Value");

                // Refresh the pivot cache and calculate data
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Save the workbook to a file
                string outputPath = "PivotTableDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
