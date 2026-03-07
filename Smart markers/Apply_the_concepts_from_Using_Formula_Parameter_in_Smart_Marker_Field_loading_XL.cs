using System;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerFormulaParameterDemo
{
    class Program
    {
        static void Main()
        {
            // Load the template workbook that contains smart markers.
            Workbook workbook = new Workbook("template.xlsx");

            // Prepare the smart marker data source (JSON string).
            string jsonData = @"{
                ""ProductName"": ""Demo Product"",
                ""Price"": 99.99,
                ""Quantity"": 10
            }";

            // Create a WorkbookDesigner and bind the workbook.
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook
            };

            // Set the JSON data source.
            designer.SetJsonDataSource("Data", jsonData);

            // Process the smart markers.
            designer.Process(true);

            // Save the processed workbook.
            workbook.Save("output.xlsx");
        }
    }
}