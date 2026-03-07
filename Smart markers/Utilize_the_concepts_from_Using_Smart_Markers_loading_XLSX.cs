using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace SmartMarkerDemo
{
    class Program
    {
        static void Main()
        {
            // Load the template workbook (XLSX) using LoadOptions (optional but follows the load rule)
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
            Workbook workbook = new Workbook("Template.xlsx", loadOptions);

            // Initialize WorkbookDesigner with the loaded workbook
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // Prepare sample data source (JSON) for smart markers
            string jsonData = @"{
                'Employees': [
                    { 'Name': 'John Doe', 'Age': 30, 'Department': 'Sales' },
                    { 'Name': 'Jane Smith', 'Age': 28, 'Department': 'HR' }
                ]
            }";

            // Bind the JSON data to a variable name used in smart markers
            designer.SetJsonDataSource("Data", jsonData);

            // Process all smart markers in the workbook
            designer.Process();

            // Save the populated workbook (follows the save rule)
            workbook.Save("Result.xlsx");
        }
    }
}