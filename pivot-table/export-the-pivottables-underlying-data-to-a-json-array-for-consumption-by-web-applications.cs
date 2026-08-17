// Title: Export PivotTable source data to a JSON array with Aspose.Cells for .NET (C#)
// Description: This example creates a workbook, fills a sheet with Category and Amount rows, builds a PivotTable on that range, and then uses Aspose.Cells JsonUtility with JsonSaveOptions (header row, empty cells, numeric values) to serialize the original source range (A1:B5) into a JSON string ready for web‑application consumption.
// Keywords: Aspose.Cells | C# export range to JSON | PivotTable JSON export | JsonUtility ExportRangeToJson | JsonSaveOptions | Excel to JSON .NET | Aspose.Cells PivotTable | web API JSON payload | REST API Excel data | Aspose.Cells example
// Common Searches: export Excel range to JSON Aspose.Cells C# | pivot table source data JSON Aspose.Cells | JsonUtility ExportRangeToJson sample | convert worksheet cells to JSON .NET | Aspose.Cells JsonSaveOptions tutorial
// Developer Intent: Generate a JSON array from the data that underlies a PivotTable for use in web or API scenarios.
// Use Cases: Return PivotTable source data as JSON from a .NET REST service. | Feed client‑side JavaScript charts with a JSON payload derived from Excel data. | Persist Excel‑derived records in a NoSQL store by serializing the source range to JSON.
// AI Prompts: Write C# code that uses Aspose.Cells to export a worksheet range to JSON, preserving the header row and empty cells. | Show how to configure JsonSaveOptions so numeric values are emitted as numbers, not strings, with Aspose.Cells. | Explain how to obtain the source range of a PivotTable and convert it to a JSON string using Aspose.Cells JsonUtility.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot; // Required for PivotTable and PivotFieldType
using Aspose.Cells.Utility;

// This example creates a workbook, fills a sheet with Category and Amount rows, builds a PivotTable on that range, and then uses Aspose.Cells JsonUtility with JsonSaveOptions (header row, empty cells, numeric values) to serialize the original source range (A1:B5) into a JSON string ready for web‑application consumption.
class ExportPivotTableDataToJson
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add sample source data
            Workbook workbook = new Workbook();
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Data";

            // Populate the source data (including a header row)
            dataSheet.Cells["A1"].PutValue("Category");
            dataSheet.Cells["B1"].PutValue("Amount");
            dataSheet.Cells["A2"].PutValue("Food");
            dataSheet.Cells["B2"].PutValue(120);
            dataSheet.Cells["A3"].PutValue("Transport");
            dataSheet.Cells["B3"].PutValue(80);
            dataSheet.Cells["A4"].PutValue("Food");
            dataSheet.Cells["B4"].PutValue(150);
            dataSheet.Cells["A5"].PutValue("Transport");
            dataSheet.Cells["B5"].PutValue(70);

            // Add a worksheet that will contain the PivotTable
            Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");

            // Create a PivotTable based on the source range A1:B5
            int pivotIndex = pivotSheet.PivotTables.Add("=Data!A1:B5", "A1", "PivotTable1");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

            // Configure the PivotTable (row field and data field)
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");
            pivotTable.ShowInTabularForm();

            // Refresh data and recalculate the PivotTable
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Get the underlying source range (the same range used to build the PivotTable)
            Aspose.Cells.Range sourceRange = workbook.Worksheets["Data"].Cells.CreateRange("A1:B5");

            // Set JSON export options (header row present, include empty cells, etc.)
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                HasHeaderRow = true,
                ExportEmptyCells = true,
                ExportAsString = false
            };

            // Convert the source range to a JSON string
            string jsonResult = JsonUtility.ExportRangeToJson(sourceRange, jsonOptions);

            // Output the JSON string (can be sent to a web application)
            Console.WriteLine(jsonResult);
        }
        catch (Exception ex)
        {
            // Log or display the error details
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
