// Title: Validate that all smart markers are resolved after processing a workbook with Aspose.Cells in C#
// AI Prompts: After calling WorkbookDesigner.Process, invoke GetSmartMarkers and verify that the returned array is empty. | If any markers are returned by GetSmartMarkers, iterate through the array and log each unresolved marker; otherwise log a success message.
// Common Searches: how to detect unresolved smart markers after WorkbookDesigner.Process in C# | Aspose.Cells get list of smart markers that were not resolved | C# verify all smart markers are replaced in an Excel template | using GetSmartMarkers to validate smart marker processing in Aspose.Cells | check for missing data fields in Aspose.Cells smart markers
// Tags: WorkbookDesigner GetSmartMarkers validation | Aspose.Cells smart marker resolution check | C# verify unresolved smart markers | Excel template smart marker processing | detect missing smart marker data Aspose.Cells

using System;
using System.Data;
using Aspose.Cells;

namespace SmartMarkerValidationDemo
{
    // The example creates an Excel workbook with smart markers, processes them using WorkbookDesigner with a DataTable data source, retrieves any remaining markers via GetSmartMarkers, validates that none remain, outputs the result, and saves the processed file.
    class Program
    {
        static void Main()
        {
            // Create a workbook that contains smart markers
            Workbook templateWorkbook = new Workbook();
            Worksheet sheet = templateWorkbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Add smart markers to the template (these will be resolved after processing)
            cells["A1"].PutValue("&=Employees.Name");
            cells["A2"].PutValue("&=Employees.Age");
            // Add an intentional unresolved marker for demonstration
            cells["A3"].PutValue("&=Employees.UnknownColumn");

            // Initialize WorkbookDesigner with the template workbook
            WorkbookDesigner designer = new WorkbookDesigner();
            designer.Workbook = templateWorkbook;

            // Prepare a simple data source
            DataTable employeeTable = new DataTable("Employees");
            employeeTable.Columns.Add("Name", typeof(string));
            employeeTable.Columns.Add("Age", typeof(int));
            employeeTable.Rows.Add("John Doe", 30);
            employeeTable.Rows.Add("Jane Smith", 28);

            // Set the data source and process the smart markers
            designer.SetDataSource(employeeTable);
            designer.Process();

            // After processing, retrieve any remaining smart markers
            string[] unresolvedMarkers = designer.GetSmartMarkers();

            // Validate that all smart markers have been resolved
            if (unresolvedMarkers.Length == 0)
            {
                Console.WriteLine("All smart markers have been successfully resolved.");
            }
            else
            {
                Console.WriteLine("Unresolved smart markers found:");
                foreach (string marker in unresolvedMarkers)
                {
                    Console.WriteLine(marker);
                }
            }

            // Save the processed workbook (optional)
            designer.Workbook.Save("ProcessedWorkbook.xlsx");
        }
    }
}
