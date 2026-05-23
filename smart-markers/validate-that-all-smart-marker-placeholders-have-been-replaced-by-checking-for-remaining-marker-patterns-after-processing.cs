using System;
using System.Data;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a workbook that contains smart markers
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("&=Employees.Name");
        sheet.Cells["A2"].PutValue("&=Employees.Age");

        // Prepare a data source matching the smart markers
        DataTable employeeTable = new DataTable("Employees");
        employeeTable.Columns.Add("Name", typeof(string));
        employeeTable.Columns.Add("Age", typeof(int));
        employeeTable.Rows.Add("John Doe", 30);
        employeeTable.Rows.Add("Jane Smith", 28);

        // Set up the WorkbookDesigner with the workbook and data source
        WorkbookDesigner designer = new WorkbookDesigner();
        designer.Workbook = workbook;
        designer.SetDataSource(employeeTable);

        // Process the smart markers – this should replace all placeholders
        designer.Process();

        // Retrieve any smart markers that remain after processing
        string[] remainingMarkers = designer.GetSmartMarkers();

        // Validate that no markers are left and report the result
        if (remainingMarkers.Length == 0)
        {
            Console.WriteLine("All smart markers have been successfully replaced.");
        }
        else
        {
            Console.WriteLine("Unreplaced smart markers detected:");
            foreach (string marker in remainingMarkers)
            {
                Console.WriteLine(marker);
            }
        }

        // Save the processed workbook
        workbook.Save("ProcessedResult.xlsx");
    }
}