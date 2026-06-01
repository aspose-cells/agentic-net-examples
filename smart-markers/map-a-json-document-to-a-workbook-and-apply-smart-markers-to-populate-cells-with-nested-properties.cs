using System;
using Aspose.Cells;

namespace AsposeCellsJsonSmartMarkerDemo
{
    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook (lifecycle create)
            Workbook workbook = new Workbook();

            // 2. Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // 3. Insert smart markers that reference nested JSON properties
            //    The syntax "&=$Data.Path.To.Property" is used for JSON data sources.
            sheet.Cells["A1"].PutValue("&=$Data.Employee.Name");
            sheet.Cells["A2"].PutValue("&=$Data.Employee.Age");
            sheet.Cells["A3"].PutValue("&=$Data.Employee.Address.City");
            sheet.Cells["A4"].PutValue("&=$Data.Employee.Address.Zip");

            // 4. Prepare a JSON string with nested objects
            string json = @"
            {
                ""Employee"": {
                    ""Name"": ""John Doe"",
                    ""Age"": 30,
                    ""Address"": {
                        ""City"": ""New York"",
                        ""Zip"": ""10001""
                    }
                }
            }";

            // 5. Create a WorkbookDesigner and assign the workbook (lifecycle load)
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook
            };

            // 6. Set the JSON string as a data source named "Data"
            designer.SetJsonDataSource("Data", json);

            // 7. Process the smart markers to populate the cells (lifecycle process)
            designer.Process();

            // 8. Save the populated workbook to disk (lifecycle save)
            workbook.Save("JsonSmartMarkerOutput.xlsx");
        }
    }
}