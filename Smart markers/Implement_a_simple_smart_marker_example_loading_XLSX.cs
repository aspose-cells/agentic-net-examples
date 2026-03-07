using System;
using Aspose.Cells;

namespace SmartMarkerExample
{
    public class SimpleSmartMarker
    {
        public static void Run()
        {
            // Load the template workbook that contains smart markers.
            Workbook templateWorkbook = new Workbook("template.xlsx");

            // Initialize the WorkbookDesigner with the loaded workbook.
            WorkbookDesigner designer = new WorkbookDesigner(templateWorkbook);

            // JSON data source matching the smart marker names.
            string jsonData = @"{
                ""Name"": ""John Doe"",
                ""Age"": 30,
                ""City"": ""New York""
            }";

            // Set the JSON data source. "Data" corresponds to the marker prefix.
            designer.SetJsonDataSource("Data", jsonData);

            // Process all smart markers.
            designer.Process();

            // Save the populated workbook.
            designer.Workbook.Save("output.xlsx");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            SimpleSmartMarker.Run();
        }
    }
}