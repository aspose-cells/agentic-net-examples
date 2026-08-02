// Title: Create a Pivot Table on a New Worksheet with Aspose.Cells for .NET (C#)
// Description: This C# example shows how to build a workbook, fill a source sheet with sample sales data, add a second worksheet, define the source range, insert a pivot table named "SalesPivot" at A1, assign Category to rows, Product to columns, and Sales to data, display it in tabular form, refresh and calculate the pivot, then save the file as PivotTableDemo.xlsx.
// Keywords: Aspose.Cells | C# pivot table | create pivot table programmatically | new worksheet pivot Aspose | source range pivot Aspose.Cells | RefreshData Aspose.Cells | CalculateData Aspose.Cells | Excel automation .NET | Aspose.Cells.Pivot example | save workbook with pivot
// Common Searches: how to add a pivot table to a new worksheet using Aspose.Cells | Aspose.Cells C# pivot table source range | set row, column, data fields in Aspose.Cells pivot table | refresh and calculate pivot table programmatically | save Excel file with pivot table Aspose
// Developer Intent: Generate a pivot table on a separate worksheet from existing data and persist the workbook.
// Use Cases: Create a sales‑summary pivot that groups totals by Category and Product for automated reporting. | Batch‑process large data sets to produce ready‑to‑use analytical worksheets without manual Excel steps. | Distribute a template workbook containing a pre‑configured pivot table that end users can refresh with their own data.
// AI Prompts: Write C# code with Aspose.Cells that creates a pivot table on a new sheet from a given data range, configures row, column, and data fields, and saves the workbook. | Explain how to change the aggregation function of the Sales field to Average and adjust the source range in the provided Aspose.Cells pivot example.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsPivotExample
{
    // This C# example shows how to build a workbook, fill a source sheet with sample sales data, add a second worksheet, define the source range, insert a pivot table named "SalesPivot" at A1, assign Category to rows, Product to columns, and Sales to data, display it in tabular form, refresh and calculate the pivot, then save the file as PivotTableDemo.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet (source data)
                Worksheet sourceSheet = workbook.Worksheets[0];
                sourceSheet.Name = "SourceData";

                // Populate sample source data
                Cells srcCells = sourceSheet.Cells;
                srcCells["A1"].PutValue("Category");
                srcCells["B1"].PutValue("Product");
                srcCells["C1"].PutValue("Sales");

                srcCells["A2"].PutValue("Fruit");
                srcCells["B2"].PutValue("Apple");
                srcCells["C2"].PutValue(1200);

                srcCells["A3"].PutValue("Fruit");
                srcCells["B3"].PutValue("Banana");
                srcCells["C3"].PutValue(850);

                srcCells["A4"].PutValue("Vegetable");
                srcCells["B4"].PutValue("Carrot");
                srcCells["C4"].PutValue(640);

                srcCells["A5"].PutValue("Vegetable");
                srcCells["B5"].PutValue("Broccoli");
                srcCells["C5"].PutValue(720);

                // Add a new worksheet that will contain the pivot table
                Worksheet pivotSheet = workbook.Worksheets.Add("PivotTable");

                // Determine the source data range address (including the header row)
                AsposeRange srcRange = srcCells.MaxDisplayRange; // e.g., A1:C5
                string sourceData = $"=SourceData!{srcRange.Address}";

                // Add a pivot table to the new sheet
                // Parameters: source data range, destination cell (upper‑left corner), pivot table name
                int pivotIndex = pivotSheet.PivotTables.Add(sourceData, "A1", "SalesPivot");

                // Retrieve the created pivot table
                PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

                // Configure the pivot table fields
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Column, "Product");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Optional: display the pivot table in tabular form
                pivotTable.ShowInTabularForm();

                // Refresh and calculate the pivot table so that data appears in the sheet
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Ensure output directory exists
                string outputPath = "PivotTableDemo.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook to a file
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
