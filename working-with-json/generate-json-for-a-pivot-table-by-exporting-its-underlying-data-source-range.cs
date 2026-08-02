using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Utility;
using Aspose.Cells.Json; // Ensure JsonSaveOptions namespace
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsPivotJsonExport
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet (data sheet)
                Workbook workbook = new Workbook();
                Worksheet dataSheet = workbook.Worksheets[0];
                dataSheet.Name = "Data";

                // Populate sample data for the pivot table source
                dataSheet.Cells["A1"].PutValue("Product");
                dataSheet.Cells["B1"].PutValue("Region");
                dataSheet.Cells["C1"].PutValue("Sales");
                dataSheet.Cells["A2"].PutValue("Product1");
                dataSheet.Cells["B2"].PutValue("North");
                dataSheet.Cells["C2"].PutValue(1000);
                dataSheet.Cells["A3"].PutValue("Product2");
                dataSheet.Cells["B3"].PutValue("South");
                dataSheet.Cells["C3"].PutValue(2000);
                dataSheet.Cells["A4"].PutValue("Product1");
                dataSheet.Cells["B4"].PutValue("East");
                dataSheet.Cells["C4"].PutValue(1500);

                // Add a new worksheet to host the pivot table
                Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");

                // Create a pivot table using the data range from the data sheet
                // Note: source string must be prefixed with the sheet name and an '=' sign
                string sourceData = $"=Data!A1:C4";
                int pivotIndex = pivotSheet.PivotTables.Add(sourceData, "A1", "SalesPivot");
                PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

                // Configure the pivot table fields
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
                pivotTable.AddFieldToArea(PivotFieldType.Column, "Region");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Populate the pivot table with calculated data
                pivotTable.CalculateData();

                // Retrieve the underlying data source range of the pivot table
                // GetSource returns an array of strings; the first element contains the range address
                string[] sourceRangeStrings = pivotTable.GetSource();
                if (sourceRangeStrings.Length == 0)
                {
                    Console.WriteLine("Pivot table source not found.");
                    return;
                }

                // Create a Range object from the source address (e.g., "A1:C4")
                AsposeRange sourceRange = dataSheet.Cells.CreateRange(sourceRangeStrings[0]);

                // Set JSON export options
                JsonSaveOptions jsonOptions = new JsonSaveOptions
                {
                    HasHeaderRow = true,          // First row contains column names
                    ExportEmptyCells = true,      // Include empty cells in the output
                    ExportNestedStructure = false // Flat JSON array
                };

                // Export the range to a JSON string
                string jsonResult = JsonUtility.ExportRangeToJson(sourceRange, jsonOptions);

                // Output the JSON string
                Console.WriteLine("Exported JSON:");
                Console.WriteLine(jsonResult);

                // Optionally, save the workbook (not required for JSON export)
                string outputPath = "PivotTableWithJsonExport.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}