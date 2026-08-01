// Title: Use anonymous .NET objects with IF conditional smart markers in Aspose.Cells (C#)
// Description: Demonstrates creating a workbook, inserting smart markers that use an IF expression to show a person's name only when Age > 30 (otherwise “Young”), displaying a boolean IsActive field, binding a List<object> of anonymous objects as the “Person” data source via WorkbookDesigner, processing the markers, and saving the result as an Excel file.
// Keywords: Aspose.Cells | C# smart markers | anonymous objects data source | IF conditional smart marker | WorkbookDesigner | conditional Excel export | boolean smart marker | list<object | dynamic data source | Excel report generation
// Common Searches: Aspose.Cells smart marker IF condition | set anonymous list as data source Aspose.Cells | conditional smart markers C# example | display boolean value with smart markers | use WorkbookDesigner with anonymous objects | dynamic data source for smart markers
// Developer Intent: Bind anonymous .NET objects to a smart‑marker data source and apply IF logic in the generated Excel file.
// Use Cases: Create age‑based personnel reports where names appear only for employees over a threshold. | Generate Excel dashboards that show active/inactive status directly from a dynamic object list. | Produce quick Excel exports from ad‑hoc anonymous collections without defining POCO classes. | Implement conditional placeholders (e.g., “Young”) in templates for marketing lists.
// AI Prompts: Show how to add multiple IF‑based smart markers using different fields from an anonymous data source. | Provide a nested IF example for smart markers with anonymous objects in Aspose.Cells. | Explain how to format a boolean smart marker as Yes/No or custom text. | Demonstrate grouping and subtotaling data when the source is a List<object> of anonymous types. | Give guidance on handling null values in conditional smart markers with anonymous objects.

using System.Collections.Generic;
using Aspose.Cells;

// Demonstrates creating a workbook, inserting smart markers that use an IF expression to show a person's name only when Age > 30 (otherwise “Young”), displaying a boolean IsActive field, binding a List<object> of anonymous objects as the “Person” data source via WorkbookDesigner, processing the markers, and saving the result as an Excel file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];

        // Add smart markers with conditional logic:
        // If Age > 30 show Name, otherwise show "Young"
        ws.Cells["A1"].PutValue("&=IF($Person.Age>30,$Person.Name,\"Young\")");
        // Show the boolean IsActive value
        ws.Cells["B1"].PutValue("&=$Person.IsActive");

        // Prepare anonymous objects as data source
        var persons = new List<object>
        {
            new { Name = "John Doe", Age = 35, IsActive = true },
            new { Name = "Jane Smith", Age = 28, IsActive = false }
        };

        // Initialize WorkbookDesigner and assign the workbook
        WorkbookDesigner designer = new WorkbookDesigner();
        designer.Workbook = wb;

        // Set the anonymous list as a data source named "Person"
        designer.SetDataSource("Person", persons);

        // Process the smart markers
        designer.Process();

        // Save the resulting workbook
        wb.Save("ConditionalSmartMarkers.xlsx");
    }
}
