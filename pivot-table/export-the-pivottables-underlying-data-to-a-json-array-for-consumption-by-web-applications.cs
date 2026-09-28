// Title: Export Aspose.Cells PivotTable source range to a formatted JSON array using C#
// AI Prompts: Write C# code that creates a workbook, adds sample data, builds a pivot table, and calls Range.ToJson with appropriate options to generate a JSON array. | Show how to set JSON export options (header row, empty‑cell handling, data type preservation, indentation) for converting an Excel range to JSON with Aspose.Cells.
// Common Searches: aspnet export pivot table source data to JSON with Aspose.Cells | c# convert excel range to json array including column headers using Aspose.Cells | how to include empty cells as null in Aspose.Cells JSON export | set indentation for JSON result from Excel in .NET | sample code for exporting pivot table data to JSON for web applications
// Tags: Aspose.Cells export range to JSON | pivot table source data JSON conversion C# | configure JSON output settings Aspose.Cells | pretty JSON output from Excel using Aspose.Cells | include null for empty cells in JSON export

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Utility;

namespace AsposeCellsPivotJsonExport
{
    // The example creates a workbook with sample data, adds a pivot table, and uses Aspose.Cells Range.ToJson with JsonSaveOptions (header row, export empty cells as null, preserve data types, indentation) to output the source range as a nicely formatted JSON array suitable for web applications.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet (source data)
                Workbook workbook = new Workbook();
                Worksheet dataSheet = workbook.Worksheets[0];
                dataSheet.Name = "Data";

                // Populate sample data for the pivot table
                dataSheet.Cells["A1"].PutValue("Category");
                dataSheet.Cells["B1"].PutValue("Product");
                dataSheet.Cells["C1"].PutValue("Quantity");

                dataSheet.Cells["A2"].PutValue("Fruit");
                dataSheet.Cells["B2"].PutValue("Apple");
                dataSheet.Cells["C2"].PutValue(120);

                dataSheet.Cells["A3"].PutValue("Fruit");
                dataSheet.Cells["B3"].PutValue("Banana");
                dataSheet.Cells["C3"].PutValue(85);

                dataSheet.Cells["A4"].PutValue("Vegetable");
                dataSheet.Cells["B4"].PutValue("Carrot");
                dataSheet.Cells["C4"].PutValue(60);

                dataSheet.Cells["A5"].PutValue("Vegetable");
                dataSheet.Cells["B5"].PutValue("Tomato");
                dataSheet.Cells["C5"].PutValue(95);

                // Add a new worksheet that will contain the pivot table
                Worksheet pivotSheet = workbook.Worksheets.Add("PivotTable");

                // Define the source range for the pivot table (A1:C5)
                string sourceRange = "Data!A1:C5";

                // Add the pivot table to the pivot sheet at cell A1
                int pivotIndex = pivotSheet.PivotTables.Add(sourceRange, "A1", "SalesPivot");
                PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

                // Configure pivot fields
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Column, "Product");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Quantity");

                // Optional: display the pivot in tabular form
                pivotTable.ShowInTabularForm();

                // Refresh and calculate the pivot table data
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Export the underlying source data (the range used for the pivot) to JSON
                Aspose.Cells.Range sourceDataRange = dataSheet.Cells.CreateRange("A1:C5");

                // Configure JSON export options
                JsonSaveOptions jsonOptions = new JsonSaveOptions
                {
                    HasHeaderRow = true,          // First row contains column names
                    ExportEmptyCells = true,      // Export empty cells as null
                    ExportAsString = false,       // Keep original data types
                    Indent = "  "                 // Pretty‑print with two‑space indent
                };

                // Convert the range to a JSON string
                string jsonResult = sourceDataRange.ToJson(jsonOptions);

                // Output the JSON string (can be sent to a web application)
                Console.WriteLine(jsonResult);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
