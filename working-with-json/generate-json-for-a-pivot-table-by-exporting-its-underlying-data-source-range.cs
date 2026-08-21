// Title: Export Pivot Table Source Range to JSON with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, add a pivot table, retrieve its source range using GetSource, build a Range object, configure JsonSaveOptions (header row, empty cells, flat structure), and export the underlying data to a JSON string with Aspose.Cells JsonUtility.
// Keywords: Aspose.Cells | C# | .NET | pivot table source range | GetSource | JsonSaveOptions | ExportRangeToJson | JSON export | range to JSON | workbook export
// Common Searches: Aspose.Cells export pivot source to JSON C# | GetSource pivot table Aspose.Cells example | JsonUtility ExportRangeToJson usage | How to export worksheet range as JSON with Aspose.Cells | C# convert pivot table source data to JSON
// Developer Intent: Extract a pivot table's underlying data range and convert it to JSON using Aspose.Cells for .NET.
// Use Cases: Provide a JSON feed of raw sales records for a web service after building a pivot table. | Create a flat JSON file that retains column headers and empty cells for downstream analytics. | Capture a reproducible JSON snapshot of the source data for auditing or version control.
// AI Prompts: Generate C# code that retrieves a pivot table's source range with GetSource and exports it to JSON using Aspose.Cells JsonUtility. | Show how to set JsonSaveOptions to include header rows and empty cells when exporting a range to JSON in Aspose.Cells. | Explain the steps to create a workbook, add a pivot table, obtain its source reference, and produce a JSON string of that data.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Utility;

namespace AsposeCellsPivotJsonExport
{
    // Demonstrates how to create a workbook, add a pivot table, retrieve its source range using GetSource, build a Range object, configure JsonSaveOptions (header row, empty cells, flat structure), and export the underlying data to a JSON string with Aspose.Cells JsonUtility.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet dataSheet = workbook.Worksheets[0];
                dataSheet.Name = "Data";

                // Populate sample data for the pivot table
                dataSheet.Cells["A1"].PutValue("Category");
                dataSheet.Cells["B1"].PutValue("Product");
                dataSheet.Cells["C1"].PutValue("Sales");

                dataSheet.Cells["A2"].PutValue("Cat1");
                dataSheet.Cells["B2"].PutValue("ProdA");
                dataSheet.Cells["C2"].PutValue(1200);

                dataSheet.Cells["A3"].PutValue("Cat1");
                dataSheet.Cells["B3"].PutValue("ProdB");
                dataSheet.Cells["C3"].PutValue(800);

                dataSheet.Cells["A4"].PutValue("Cat2");
                dataSheet.Cells["B4"].PutValue("ProdC");
                dataSheet.Cells["C4"].PutValue(1500);

                // Add a worksheet that will contain the pivot table
                Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");

                // Define the source data range address (including the header row)
                string sourceData = $"=Data!{dataSheet.Cells.MaxDisplayRange.Address}";

                // Add the pivot table to the pivot sheet
                int pivotIndex = pivotSheet.PivotTables.Add(sourceData, "A1", "SalesPivot");
                PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

                // Configure pivot fields (Category as row, Sales as data)
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Refresh and calculate the pivot table using the correct API
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Retrieve the underlying data source reference (e.g., "A1:C4")
                string[] sourceRefs = pivotTable.GetSource();
                if (sourceRefs.Length == 0)
                {
                    Console.WriteLine("Pivot table source not found.");
                    return;
                }

                // Create a Range object based on the source reference
                Aspose.Cells.Range sourceRange = dataSheet.Cells.CreateRange(sourceRefs[0]);

                // Set JSON export options
                JsonSaveOptions jsonOptions = new JsonSaveOptions
                {
                    HasHeaderRow = true,          // First row contains column names
                    ExportEmptyCells = true,      // Include empty cells in the output
                    ExportNestedStructure = false // Flat JSON array
                };

                // Export the source range to a JSON string
                string json = JsonUtility.ExportRangeToJson(sourceRange, jsonOptions);

                // Output the JSON string
                Console.WriteLine("Exported JSON:");
                Console.WriteLine(json);

                // Save the workbook
                workbook.Save("PivotTableWithJsonExport.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
