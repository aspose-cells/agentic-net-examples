using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Utility;

class ExportPivotSourceToJson
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Data";

            dataSheet.Cells["A1"].PutValue("Category");
            dataSheet.Cells["B1"].PutValue("Amount");
            dataSheet.Cells["A2"].PutValue("Food");
            dataSheet.Cells["B2"].PutValue(100);
            dataSheet.Cells["A3"].PutValue("Drink");
            dataSheet.Cells["B3"].PutValue(150);
            dataSheet.Cells["A4"].PutValue("Food");
            dataSheet.Cells["B4"].PutValue(200);

            // Add a worksheet for the pivot table
            Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");

            // Create a pivot table based on the data range
            string sourceData = "Data!A1:B4";
            int pivotIndex = pivotSheet.PivotTables.Add(sourceData, "A1", "MyPivot");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");
            pivotTable.CalculateData();

            // Retrieve the underlying data source range string from the pivot table
            string[] sourceArray = pivotTable.GetSource(); // e.g., ["A1:B4"]
            if (sourceArray.Length > 0)
            {
                // Build a Range object on the original data worksheet
                Aspose.Cells.Range dataRange = dataSheet.Cells.CreateRange(sourceArray[0]);

                // Set JSON export options
                JsonSaveOptions jsonOptions = new JsonSaveOptions
                {
                    HasHeaderRow = true,
                    ExportEmptyCells = true
                };

                // Export the range to a JSON string
                string json = JsonUtility.ExportRangeToJson(dataRange, jsonOptions);

                // Output the JSON result
                Console.WriteLine(json);
            }

            // Save the workbook (optional, not required for JSON export)
            string outputPath = "PivotExportDemo.xlsx";
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}