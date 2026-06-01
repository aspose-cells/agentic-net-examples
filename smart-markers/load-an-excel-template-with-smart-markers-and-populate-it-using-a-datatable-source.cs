using System;
using System.Data;
using Aspose.Cells;

class SmartMarkerExample
{
    static void Main()
    {
        // Load the Excel template that contains smart markers (e.g., &=$Employees.Name)
        Workbook template = new Workbook("TemplateWithSmartMarkers.xlsx");

        // Create a DataTable that matches the smart marker table name and columns
        DataTable employees = new DataTable("Employees");
        employees.Columns.Add("Name", typeof(string));
        employees.Columns.Add("Age", typeof(int));
        employees.Columns.Add("Department", typeof(string));

        // Populate the DataTable with sample data
        employees.Rows.Add("John Doe", 30, "Sales");
        employees.Rows.Add("Jane Smith", 28, "Marketing");
        employees.Rows.Add("Mike Johnson", 35, "Engineering");

        // Initialize WorkbookDesigner with the loaded template workbook
        WorkbookDesigner designer = new WorkbookDesigner(template);

        // Bind the DataTable as the data source for the smart markers
        designer.SetDataSource(employees);

        // Process the smart markers and fill the worksheet with data
        designer.Process();

        // Save the resulting workbook
        designer.Workbook.Save("PopulatedOutput.xlsx");
    }
}