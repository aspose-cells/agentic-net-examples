// Title: Use Aspose.Cells WorkbookDesigner to bind a JSON string and populate smart markers in a C# Excel report
// AI Prompts: Generate C# code that creates a Workbook, adds smart markers referencing JSON fields, assigns the workbook to a WorkbookDesigner, sets a JSON data source with SetJsonDataSource, processes the markers, and saves the file. | Show how to bind a JSON string as a named data source to WorkbookDesigner and fill an Excel template using smart markers in Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# WorkbookDesigner SetJsonDataSource example with smart markers | populate Excel template from JSON string using Aspose.Cells smart markers | how to bind JSON data to smart markers in Aspose.Cells for .NET | C# generate Excel report from JSON using WorkbookDesigner
// Tags: SetJsonDataSource JSON binding | Aspose.Cells smart markers Excel population | C# create Excel file from JSON data | process smart markers with WorkbookDesigner | Aspose.Cells JSON-driven Excel report generation

using System;
using Aspose.Cells;

// The example creates a new Workbook, inserts smart markers that reference Employee fields, assigns the workbook to a WorkbookDesigner, sets a JSON string as the "Employee" data source via SetJsonDataSource, processes the smart markers to fill the cells, and saves the populated workbook as EmployeeReport.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Add smart markers that map to JSON fields
        sheet.Cells["A1"].PutValue("&=$Employee.Name");
        sheet.Cells["B1"].PutValue("&=$Employee.Age");
        sheet.Cells["C1"].PutValue("&=$Employee.City");

        // Initialize WorkbookDesigner and assign the workbook
        WorkbookDesigner designer = new WorkbookDesigner();
        designer.Workbook = workbook;

        // JSON string representing the data source
        string json = "{\"Name\":\"John Doe\",\"Age\":30,\"City\":\"New York\"}";

        // Set the JSON data source; the first parameter is the data source name
        designer.SetJsonDataSource("Employee", json);

        // Process the smart markers to populate the worksheet
        designer.Process();

        // Save the populated workbook
        workbook.Save("EmployeeReport.xlsx");
    }
}
