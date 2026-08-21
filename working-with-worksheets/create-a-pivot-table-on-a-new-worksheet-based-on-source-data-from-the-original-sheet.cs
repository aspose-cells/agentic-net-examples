// Title: Create a Pivot Table on a New Worksheet with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to generate a workbook, fill the first sheet with sample data, compute the used range, add a second sheet, insert a pivot table that references the source range, assign "Category" as a row field and "Amount" as a data field, apply a medium‑style theme, refresh the pivot, and save the file as PivotTableDemo.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells pivot table C# | add pivot table new worksheet .NET | source range MaxDisplayRange Aspose.Cells | refresh pivot tables Aspose.Cells | pivot table style Aspose.Cells | create workbook Aspose.Cells | C# Excel pivot example
// Common Searches: how to add a pivot table on a separate sheet using Aspose.Cells | Aspose.Cells example creating pivot table from existing worksheet | set row and data fields in Aspose.Cells pivot table | save workbook with pivot table Aspose.Cells C# | determine source range for pivot table Aspose.Cells
// Developer Intent: Generate a pivot table on a newly added worksheet that summarizes data from the original sheet.
// Use Cases: Produce a sales‑by‑category summary on a dedicated report sheet. | Automate financial pivot tables across many workbooks with a uniform style. | Refresh all pivots after bulk data updates before exporting to Excel.
// AI Prompts: Write C# code that creates a pivot table on a new worksheet with Aspose.Cells, uses "Category" as the row field and "Amount" as the data field, applies a medium style, refreshes the pivot, and saves the workbook. | Explain how to programmatically obtain the used range of a worksheet and pass it to PivotTables.Add in Aspose.Cells. | Show the steps to refresh every pivot table in a workbook after modifying source data using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotDemo
{
    // Helper class that creates a pivot table on a new worksheet
    // Demonstrates how to generate a workbook, fill the first sheet with sample data, compute the used range, add a second sheet, insert a pivot table that references the source range, assign "Category" as a row field and "Amount" as a data field, apply a medium‑style theme, refresh the pivot, and save the file as PivotTableDemo.xlsx using Aspose.Cells for .NET.
    public class CreatePivotOnNewSheet
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the original worksheet (source data) and name it
                Worksheet sourceSheet = workbook.Worksheets[0];
                sourceSheet.Name = "SourceData";

                // Populate sample source data
                sourceSheet.Cells["A1"].PutValue("Category");
                sourceSheet.Cells["B1"].PutValue("Amount");
                sourceSheet.Cells["A2"].PutValue("Fruit");
                sourceSheet.Cells["B2"].PutValue(120);
                sourceSheet.Cells["A3"].PutValue("Vegetable");
                sourceSheet.Cells["B3"].PutValue(80);
                sourceSheet.Cells["A4"].PutValue("Fruit");
                sourceSheet.Cells["B4"].PutValue(150);
                sourceSheet.Cells["A5"].PutValue("Vegetable");
                sourceSheet.Cells["B5"].PutValue(70);

                // Determine the source data range (e.g., A1:B5)
                Aspose.Cells.Range sourceRange = sourceSheet.Cells.MaxDisplayRange;
                string sourceData = $"=SourceData!{sourceRange.Address}";

                // Add a new worksheet to host the pivot table
                Worksheet pivotSheet = workbook.Worksheets.Add("PivotTable");

                // Add a pivot table to the new sheet using the source data
                // Destination cell is A1 in the pivot sheet, table name is "MyPivot"
                int pivotIndex = pivotSheet.PivotTables.Add(sourceData, "A1", "MyPivot");

                // Retrieve the created pivot table
                PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

                // Configure the pivot table: Category as row field, Amount as data field
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

                // Optional: set a style and refresh the pivot table
                pivotTable.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;
                pivotSheet.RefreshPivotTables();

                // Save the workbook
                workbook.Save("PivotTableDemo.xlsx");
                Console.WriteLine("Pivot table created and saved as PivotTableDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    class Program
    {
        static void Main(string[] args)
        {
            CreatePivotOnNewSheet.Run();
        }
    }
}
