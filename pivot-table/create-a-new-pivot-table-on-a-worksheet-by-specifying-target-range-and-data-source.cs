// Title: Create a Pivot Table on a Separate Worksheet by Specifying Source Range and Target Cell Using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a new Workbook, populates it with sample data, builds a source range reference string (e.g., =SourceData!A1:B4), and inserts a PivotTable named "MyPivotTable" at cell A1 on a newly added worksheet using Aspose.Cells. | Show how to configure the PivotTable by adding a row field for "Category" and a data field for "Value", calculate the pivot data, and save the workbook as an .xlsx file.
// Common Searches: Aspose.Cells C# add pivot table to a new worksheet with source range reference | How to set pivot table source data string in Aspose.Cells .NET | Create pivot table programmatically in C# using Aspose.Cells and calculate its data | Aspose.Cells example defining target cell for pivot table insertion
// Tags: Aspose.Cells programmatic pivot table insertion | Aspose.Cells define pivot source range string | Aspose.Cells configure pivot fields C# | Aspose.Cells calculate pivot data | Aspose.Cells save workbook with pivot table

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using AsposeRange = Aspose.Cells.Range;

// Demonstrates creating a workbook, adding sample data, defining a source range reference, inserting a PivotTable on a separate worksheet at cell A1, configuring row and data fields, calculating the pivot results, and saving the file as PivotTableDemo.xlsx using Aspose.Cells for .NET.
class CreatePivotTableDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet and name it as the source data sheet
            Worksheet sourceSheet = workbook.Worksheets[0];
            sourceSheet.Name = "SourceData";

            // Populate sample data for the pivot table
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

            // Build the source data reference string (e.g., =SourceData!A1:B4)
            AsposeRange sourceRange = sourceSheet.Cells.MaxDisplayRange;
            string sourceData = $"=SourceData!{sourceRange.Address}";

            // Add a pivot table to the pivot sheet at cell A1
            PivotTableCollection pivotTables = pivotSheet.PivotTables;
            int pivotIndex = pivotTables.Add(sourceData, "A1", "MyPivotTable");

            // Retrieve the created pivot table
            PivotTable pivotTable = pivotTables[pivotIndex];

            // Configure the pivot table: add a row field and a data field
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Value");

            // Calculate the pivot table data
            pivotTable.CalculateData();

            // Save the workbook to a file
            workbook.Save("PivotTableDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
