using System;
using System.Data;
using Aspose.Cells;
using Aspose.Cells.Markup;

class SmartMarkerArrayAccess
{
    static void Main()
    {
        // Load the template workbook that contains smart markers
        // Example smart marker in cell A1: &=Employees[1].Name  (zero‑based index)
        string templatePath = "template.xlsx";
        Workbook workbook = new Workbook(templatePath);

        // Initialize WorkbookDesigner with the loaded workbook
        WorkbookDesigner designer = new WorkbookDesigner();
        designer.Workbook = workbook;

        // Prepare a data source (DataTable) that will be used by the smart markers
        DataTable employees = new DataTable("Employees");
        employees.Columns.Add("Name", typeof(string));
        employees.Columns.Add("Age", typeof(int));

        // Add sample rows – the index used in the smart marker refers to these rows
        employees.Rows.Add("Alice", 30);   // index 0
        employees.Rows.Add("Bob", 25);     // index 1
        employees.Rows.Add("Charlie", 35); // index 2

        // Set the data source for the designer
        designer.SetDataSource(employees);

        // Process all smart markers in the workbook
        designer.Process();

        // After processing, the cell that contained the smart marker now holds the value
        // of the specific array element (zero‑based index 1 -> "Bob")
        Cell resultCell = designer.Workbook.Worksheets[0].Cells["A1"]; // using Cells indexer
        Console.WriteLine("Value from Employees[1].Name: " + resultCell.StringValue);

        // Save the processed workbook
        designer.Workbook.Save("output.xlsx");
    }
}