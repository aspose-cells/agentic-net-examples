using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Utility;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsPivotJsonExport
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet (data sheet)
            Workbook workbook = new Workbook();
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Data";

            // Populate sample data for the pivot table
            dataSheet.Cells["A1"].PutValue("Category");
            dataSheet.Cells["B1"].PutValue("Product");
            dataSheet.Cells["C1"].PutValue("Sales");
            dataSheet.Cells["A2"].PutValue("Fruit");
            dataSheet.Cells["B2"].PutValue("Apple");
            dataSheet.Cells["C2"].PutValue(1200);
            dataSheet.Cells["A3"].PutValue("Fruit");
            dataSheet.Cells["B3"].PutValue("Banana");
            dataSheet.Cells["C3"].PutValue(800);
            dataSheet.Cells["A4"].PutValue("Vegetable");
            dataSheet.Cells["B4"].PutValue("Carrot");
            dataSheet.Cells["C4"].PutValue(500);

            // Add a new worksheet that will host the pivot table
            Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");

            // Define the source data range address (A1:C4) and add the pivot table
            string sourceData = "=Data!A1:C4";
            int pivotIndex = pivotSheet.PivotTables.Add(sourceData, "A1", "SalesPivot");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

            // Configure pivot fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Column, "Product");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Refresh and calculate the pivot table so that it is populated
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Retrieve the original data source address from the pivot table
            string[] sourceInfo = pivotTable.GetSource();
            if (sourceInfo == null || sourceInfo.Length == 0)
            {
                Console.WriteLine("Pivot table source not found.");
                return;
            }

            // sourceInfo[0] contains the source range (e.g., "Data!A1:C4")
            string sourceRangeStr = sourceInfo[0];
            // Remove any leading '=' if present.
            sourceRangeStr = sourceRangeStr.TrimStart('=');
            // Extract the address part after the sheet name (if present)
            string addressPart = sourceRangeStr.Contains("!") ? sourceRangeStr.Split('!')[1] : sourceRangeStr;

            // Create a Range object that represents the underlying data source
            AsposeRange sourceRange = dataSheet.Cells.CreateRange(addressPart);

            // Set JSON export options
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                ExportEmptyCells = true,
                HasHeaderRow = true,
                ExportNestedStructure = false
            };

            // Export the range to a JSON string
            string jsonResult = JsonUtility.ExportRangeToJson(sourceRange, jsonOptions);

            // Output the JSON string
            Console.WriteLine("Exported JSON:");
            Console.WriteLine(jsonResult);

            // Save the workbook (optional)
            workbook.Save("PivotTableWithJsonExport.xlsx");
        }
    }
}