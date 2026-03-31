using System;
using Aspose.Cells;

class SmartMarkerJsonImport
{
    static void Main()
    {
        // Create a new workbook (empty template)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Insert smart markers that will be replaced by JSON data
        // The marker syntax "&=$DataSource.Property"
        sheet.Cells["A1"].PutValue("&=$Employee.Name");
        sheet.Cells["B1"].PutValue("&=$Employee.Age");
        sheet.Cells["C1"].PutValue("&=$Employee.City");

        // Initialize the WorkbookDesigner and assign the workbook
        WorkbookDesigner designer = new WorkbookDesigner();
        designer.Workbook = workbook;

        // JSON string representing the data source
        string json = "{\"Name\":\"John Doe\",\"Age\":30,\"City\":\"New York\"}";

        // Bind the JSON string to the smart marker data source named "Employee"
        designer.SetJsonDataSource("Employee", json);

        // Process the smart markers – they will be replaced with JSON values
        designer.Process();

        // Save the populated workbook
        workbook.Save("SmartMarkerJsonOutput.xlsx");
    }
}