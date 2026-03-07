using System;
using Aspose.Cells;

class SmartMarkerRangeDemo
{
    static void Main()
    {
        // Create a new workbook and add smart markers
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Place smart markers in the range A2:B2
        sheet.Cells["A2"].PutValue("&=Data.Name");
        sheet.Cells["B2"].PutValue("&=Data.Value");

        // Define the range that holds the smart markers and name it for range smart markers
        var smartRange = sheet.Cells.CreateRange("A2:B2");
        smartRange.Name = "_CellsSmartMarkers";

        // Initialize WorkbookDesigner with the workbook
        WorkbookDesigner designer = new WorkbookDesigner(workbook);

        // Provide JSON data source matching the smart markers
        string jsonData = "{\"Name\":\"Sample Product\",\"Value\":123.45}";
        designer.SetJsonDataSource("Data", jsonData);

        // Process only the specified range; true = preserve unrecognized markers
        designer.Process(smartRange, true);

        // Save the processed workbook
        workbook.Save("ProcessedOutput.xlsx");
    }
}