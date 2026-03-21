using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace SmartMarkerRangeProcessing
{
    class Program
    {
        static void Main()
        {
            // Load the template workbook that contains smart markers
            Workbook workbook = new Workbook("template.xlsx");

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Define the range that encloses the smart markers to be processed
            // Example range A2:B2 – modify according to your template
            AsposeRange smartMarkerRange = sheet.Cells.CreateRange("A2:B2");
            // Name the range with the reserved name for range smart markers
            smartMarkerRange.Name = "_CellsSmartMarkers";

            // Initialize the WorkbookDesigner and assign the loaded workbook
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook
            };

            // Provide a JSON data source that matches the smart marker expressions
            // Example smart markers: &=$Data.Name and &=$Data.Value
            string jsonData = "{\"Name\":\"Test Product\",\"Value\":100.50}";
            designer.SetJsonDataSource("Data", jsonData);

            // Process only the defined range (true = preserve unrecognized markers)
            designer.Process(smartMarkerRange, true);

            // Save the resulting workbook
            workbook.Save("output.xlsx");
        }
    }
}