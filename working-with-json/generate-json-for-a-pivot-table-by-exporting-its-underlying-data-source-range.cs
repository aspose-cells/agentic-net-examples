// Title: Export the underlying data source range of an Aspose.Cells pivot table to a JSON file using C#
// Common Searches: C# Aspose.Cells export pivot table source range as JSON file | How to retrieve a pivot table data source range and save it to JSON with Aspose.Cells | Aspose.Cells ToJson method for converting an Excel range to JSON in .NET | Export Excel range to JSON while preserving data types using JsonSaveOptions | Save pivot table underlying data to JSON using Aspose.Cells C# example
// Tags: pivot table data source range export JSON | Aspose.Cells ToJson with JsonSaveOptions | C# export Excel range as JSON | save JSON file from Excel range Aspose.Cells | retrieve pivot table source range Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Utility; // For JsonSaveOptions

namespace PivotTableJsonExport
{
    // The example creates a workbook, fills it with sample data, adds a pivot table, obtains the original source range, configures JsonSaveOptions to keep the header row and preserve data types, converts the range to a JSON string with ToJson, writes the JSON to a file, and finally saves the workbook containing the pivot table.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Product");
                sheet.Cells["C1"].PutValue("Quantity");
                sheet.Cells["D1"].PutValue("Price");

                sheet.Cells["A2"].PutValue("Fruit");
                sheet.Cells["B2"].PutValue("Apple");
                sheet.Cells["C2"].PutValue(120);
                sheet.Cells["D2"].PutValue(0.5);

                sheet.Cells["A3"].PutValue("Fruit");
                sheet.Cells["B3"].PutValue("Banana");
                sheet.Cells["C3"].PutValue(150);
                sheet.Cells["D3"].PutValue(0.3);

                sheet.Cells["A4"].PutValue("Vegetable");
                sheet.Cells["B4"].PutValue("Carrot");
                sheet.Cells["C4"].PutValue(200);
                sheet.Cells["D4"].PutValue(0.2);

                sheet.Cells["A5"].PutValue("Vegetable");
                sheet.Cells["B5"].PutValue("Tomato");
                sheet.Cells["C5"].PutValue(180);
                sheet.Cells["D5"].PutValue(0.4);

                // Define source and destination ranges
                string sourceRange = "A1:D5";
                string destination = "F3";

                // Add a pivot table (obsolete overload – kept for simplicity)
#pragma warning disable CS0618 // Suppress obsolete warning
                int pivotIndex = sheet.PivotTables.Add(sourceRange, destination, "SalesPivot");
#pragma warning restore CS0618
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Configure the pivot table fields
                pivotTable.AddFieldToArea(PivotFieldType.Row, 0);      // Category
                pivotTable.AddFieldToArea(PivotFieldType.Column, 1); // Product
                pivotTable.AddFieldToArea(PivotFieldType.Data, 2);   // Quantity
                pivotTable.AddFieldToArea(PivotFieldType.Data, 3);   // Price

                // Refresh the pivot table so that it reflects the source data
#pragma warning disable CS0618 // Suppress obsolete warning for RefreshData
                pivotTable.RefreshData();
#pragma warning restore CS0618
                pivotTable.CalculateData();

                // Obtain the data source range (same as the original source range)
                Aspose.Cells.Range dataSourceRange = sheet.Cells.CreateRange(sourceRange);
                if (dataSourceRange == null)
                {
                    Console.WriteLine("Unable to obtain the data source range from the pivot table.");
                    return;
                }

                // Set up JSON export options
                JsonSaveOptions jsonOptions = new JsonSaveOptions
                {
                    HasHeaderRow = true,   // First row contains column names
                    ExportAsString = false // Preserve original data types where possible
                };

                // Export the range to a JSON string
                string json = dataSourceRange.ToJson(jsonOptions);

                // Output the JSON to console
                Console.WriteLine("Exported JSON:");
                Console.WriteLine(json);

                // Save the JSON to a file
                string jsonFilePath = "PivotDataSource.json";
                File.WriteAllText(jsonFilePath, json);
                Console.WriteLine($"JSON has been saved to: {Path.GetFullPath(jsonFilePath)}");

                // Optionally, save the workbook to verify the pivot table creation
                string workbookPath = "PivotTableDemo.xlsx";
                workbook.Save(workbookPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook with pivot table saved as {Path.GetFullPath(workbookPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
