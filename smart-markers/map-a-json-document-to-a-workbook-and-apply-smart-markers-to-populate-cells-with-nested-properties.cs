// Title: Use Aspose.Cells Smart Markers in C# to Populate an Excel Workbook from Nested JSON Employee Data
// AI Prompts: Write C# code that creates an Aspose.Cells workbook, places smart markers for Employee.Name, Employee.Age, Employee.Address.City, and Employee.Address.ZipCode, binds a hierarchical JSON string as a data source with WorkbookDesigner, processes the markers, and saves the result as an Excel file. | Show how to configure WorkbookDesigner with a JSON data source and apply smart‑marker syntax for nested object properties to generate a populated Excel report in C#.
// Common Searches: Aspose.Cells C# example for mapping nested JSON objects to Excel using smart markers | How to bind hierarchical JSON to smart markers in an Aspose.Cells workbook | C# populate Excel cells with Employee.Address.City from JSON via WorkbookDesigner | Smart marker syntax for nested properties in Aspose.Cells C# tutorial
// Tags: Aspose.Cells WorkbookDesigner JSON data source | smart markers nested object mapping | populate Excel from JSON hierarchy C# | Excel report generation from employee JSON Aspose.Cells | C# smart marker syntax for nested properties

using System;
using Aspose.Cells;

namespace AsposeCellsJsonSmartMarkerDemo
{
    // The sample creates a new workbook, inserts smart markers that reference nested JSON fields (Employee.Name, Employee.Age, Employee.Address.City, Employee.Address.ZipCode), sets a JSON string as the data source named "Data" using WorkbookDesigner, processes the markers to replace them with actual values, and saves the populated workbook as EmployeeReport.xlsx.
    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook (template)
            Workbook workbook = new Workbook();

            // 2. Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // 3. Place smart markers that reference nested JSON properties
            //    The JSON will contain an object "Employee" with nested "Address"
            //    Marker syntax: &=$Data.Employee.Name   -> top‑level property
            //    Marker syntax: &=$Data.Employee.Address.City -> nested property
            sheet.Cells["A1"].PutValue("Name:");
            sheet.Cells["B1"].PutValue("&=$Data.Employee.Name");
            sheet.Cells["A2"].PutValue("Age:");
            sheet.Cells["B2"].PutValue("&=$Data.Employee.Age");
            sheet.Cells["A3"].PutValue("City:");
            sheet.Cells["B3"].PutValue("&=$Data.Employee.Address.City");
            sheet.Cells["A4"].PutValue("Zip:");
            sheet.Cells["B4"].PutValue("&=$Data.Employee.Address.ZipCode");

            // 4. Prepare JSON data that matches the smart marker hierarchy
            string json = @"
            {
                ""Employee"": {
                    ""Name"": ""John Doe"",
                    ""Age"": 35,
                    ""Address"": {
                        ""City"": ""New York"",
                        ""ZipCode"": ""10001""
                    }
                }
            }";

            // 5. Create a WorkbookDesigner, assign the workbook and set the JSON data source
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook
            };
            // The first parameter is the name of the data source; it can be any identifier.
            designer.SetJsonDataSource("Data", json);

            // 6. Process the smart markers – this will replace the markers with actual values
            designer.Process();

            // 7. Save the populated workbook
            workbook.Save("EmployeeReport.xlsx");
        }
    }
}
