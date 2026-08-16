// Title: C# – Create a Pivot Table from a Range on a New Worksheet with Aspose.Cells for .NET
// Description: A complete Aspose.Cells for .NET example that creates a workbook, fills a source sheet with data, defines the source range, adds a new worksheet, inserts a pivot table at A1, sets a row field (Category) and a data field (Value), refreshes the pivot, and saves the file as PivotTableDemo.xlsx.
// Keywords: Aspose.Cells C# pivot table example | create pivot table new worksheet .NET | Aspose.Cells source range pivot | RefreshPivotTables Aspose.Cells | Workbook.Save pivot table | GitHub Aspose.Cells sample | C# Excel pivot table automation | Aspose.Cells PivotTableCollection | Aspose.Cells add field to area
// Common Searches: how to add a pivot table to a new sheet using Aspose.Cells | Aspose.Cells create pivot table from range C# | refresh pivot tables Aspose.Cells .NET | sample code Aspose.Cells pivot table new worksheet | Aspose.Cells PivotTableCollection Add example
// Developer Intent: Generate a .NET workbook that contains a pivot table built from a defined range and placed on a separate worksheet.
// Use Cases: Automate monthly sales summaries by generating a pivot table on a dedicated sheet for each reporting period. | Create dynamic financial dashboards that add a pivot table to a new worksheet for each expense category. | Produce batch Excel reports where raw transaction data is transformed into pivot tables for quick analysis.
// AI Prompts: Write C# code using Aspose.Cells to create a pivot table from a specified range, place it on a new worksheet, refresh it, and save the workbook. | Show an Aspose.Cells example that adds multiple data fields to a pivot table, applies basic number formatting, and saves the result. | Explain how to change the source range of an existing Aspose.Cells pivot table programmatically and refresh the table in C#.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using AsposeRange = Aspose.Cells.Range;

// A complete Aspose.Cells for .NET example that creates a workbook, fills a source sheet with data, defines the source range, adds a new worksheet, inserts a pivot table at A1, sets a row field (Category) and a data field (Value), refreshes the pivot, and saves the file as PivotTableDemo.xlsx.
class CreatePivotTable
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet and set it as the source data sheet
            Worksheet sourceSheet = workbook.Worksheets[0];
            sourceSheet.Name = "SourceData";

            // Populate sample data
            Cells srcCells = sourceSheet.Cells;
            srcCells["A1"].PutValue("Category");
            srcCells["B1"].PutValue("Value");
            srcCells["A2"].PutValue("A");
            srcCells["B2"].PutValue(10);
            srcCells["A3"].PutValue("B");
            srcCells["B3"].PutValue(20);
            srcCells["A4"].PutValue("A");
            srcCells["B4"].PutValue(30);

            // Add a new worksheet that will contain the pivot table
            Worksheet pivotSheet = workbook.Worksheets.Add("PivotTable");

            // Build the source data reference string (e.g., =SourceData!A1:B4)
            AsposeRange srcRange = sourceSheet.Cells.MaxDisplayRange;
            string sourceData = $"=SourceData!{srcRange.Address}";

            // Add the pivot table
            PivotTableCollection pivots = pivotSheet.PivotTables;
            int pivotIdx = pivots.Add(sourceData, "A1", "MyPivotTable");

            // Configure the pivot table: add a row field and a data field
            PivotTable pivot = pivots[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, "Category");
            pivot.AddFieldToArea(PivotFieldType.Data, "Value");

            // Refresh the pivot table to populate it with data
            pivotSheet.RefreshPivotTables();

            // Save the workbook
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
