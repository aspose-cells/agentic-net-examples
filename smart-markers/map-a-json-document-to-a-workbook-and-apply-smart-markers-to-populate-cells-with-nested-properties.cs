using System;
using Aspose.Cells;

namespace JsonToWorkbookSmartMarkers
{
    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // 2. Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // 3. Insert smart markers that reference nested JSON properties
            //    The syntax "&=$Data.Path" tells Aspose.Cells to replace the marker with the value from the JSON data source.
            sheet.Cells["A1"].PutValue("&=$Data.Employee.Name");
            sheet.Cells["A2"].PutValue("&=$Data.Employee.Address.City");
            sheet.Cells["A3"].PutValue("&=$Data.Employee.Address.Zip");

            // 4. Prepare a JSON string with nested objects
            string json = @"
            {
                ""Employee"": {
                    ""Name"": ""John Doe"",
                    ""Address"": {
                        ""City"": ""New York"",
                        ""Zip"": ""10001""
                    }
                }
            }";

            // 5. Set up the WorkbookDesigner (lifecycle: create)
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook // bind the workbook to the designer
            };

            // 6. Bind the JSON string as a data source named "Data"
            //    This uses the SetJsonDataSource method as required.
            designer.SetJsonDataSource("Data", json);

            // 7. Process the smart markers to populate the cells with JSON values
            designer.Process();

            // 8. Save the populated workbook (lifecycle: save)
            workbook.Save("Output.xlsx");
        }
    }
}