// Title: Bind JSON to WorkbookDesigner and Process Smart Markers in Aspose.Cells for .NET (C#)
// Description: Shows how to create a workbook template, insert smart markers that reference an "Employee" JSON data source, assign the workbook to a WorkbookDesigner, bind a JSON string, process the markers, and save the populated Excel file (EmployeeReport.xlsx).
// Keywords: Aspose.Cells | C# | WorkbookDesigner | JSON data source | smart markers | populate Excel | Excel report generation | Aspose.Cells for .NET
// Common Searches: Aspose.Cells bind JSON to WorkbookDesigner | C# smart markers from JSON | populate Excel from JSON Aspose.Cells | WorkbookDesigner SetJsonDataSource example | process smart markers C#
// Developer Intent: Bind a JSON string to a WorkbookDesigner, process smart markers, and generate a fully populated Excel workbook.
// Use Cases: Create an employee report by mapping JSON fields (Name, Age, City) to smart markers in a template workbook. | Generate a sales summary sheet where each sale record from a JSON array populates rows via smart markers. | Build a product catalog Excel file by assigning a JSON array of product details to WorkbookDesigner and expanding smart markers.
// AI Prompts: Show me how to bind a JSON array to WorkbookDesigner and expand smart markers into multiple rows. | Explain how to handle nested JSON objects with smart markers in Aspose.Cells. | Provide code for error handling when required JSON fields are missing during WorkbookDesigner.Process().

using System;
using Aspose.Cells;

// Shows how to create a workbook template, insert smart markers that reference an "Employee" JSON data source, assign the workbook to a WorkbookDesigner, bind a JSON string, process the markers, and save the populated Excel file (EmployeeReport.xlsx).
class Program
{
    static void Main()
    {
        // Create a new workbook (template)
        Workbook workbook = new Workbook();

        // Get the first worksheet where smart markers will be placed
        Worksheet sheet = workbook.Worksheets[0];

        // Insert smart markers that reference fields from the JSON data source
        sheet.Cells["A1"].PutValue("&=$Employee.Name");
        sheet.Cells["B1"].PutValue("&=$Employee.Age");
        sheet.Cells["C1"].PutValue("&=$Employee.City");

        // Initialize the WorkbookDesigner and assign the workbook to it
        WorkbookDesigner designer = new WorkbookDesigner();
        designer.Workbook = workbook;

        // JSON string that will serve as the data source
        string json = "{\"Name\":\"John Doe\",\"Age\":30,\"City\":\"New York\"}";

        // Bind the JSON string to the data source name used in the smart markers
        designer.SetJsonDataSource("Employee", json);

        // Process the smart markers and populate the worksheet with JSON data
        designer.Process();

        // Save the resulting workbook
        workbook.Save("EmployeeReport.xlsx");
    }
}
