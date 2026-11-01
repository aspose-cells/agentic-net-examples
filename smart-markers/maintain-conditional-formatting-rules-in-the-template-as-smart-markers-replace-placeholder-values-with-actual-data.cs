// Title: Keep existing conditional formatting when filling smart markers in an Excel template with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an Excel template containing conditional formatting, binds a DataTable to smart markers using WorkbookDesigner, processes the markers, and saves the workbook while preserving all conditional formatting rules. | Show how to use Aspose.Cells WorkbookDesigner to populate smart markers from a DataTable without modifying the template's conditional formatting in a .NET application.
// Common Searches: asp.net how to retain conditional formatting after processing smart markers with Aspose.Cells | c# Aspose.Cells WorkbookDesigner keep conditional formatting rules | populate Excel smart markers from DataTable without losing formatting | preserve conditional formatting in Excel template when using Aspose.Cells smart markers
// Tags: WorkbookDesigner preserve formatting | C# bind DataTable to smart markers | Excel template conditional formatting Aspose.Cells | process smart markers without altering formatting

using System;
using System.Data;
using Aspose.Cells;
using Aspose.Cells.Tables;

// Loads an Excel template with smart markers and conditional formatting, binds a DataTable as the data source, processes the smart markers using WorkbookDesigner, and saves the workbook while keeping the original conditional formatting rules intact.
class ConditionalFormattingWithSmartMarkers
{
    static void Main()
    {
        // Load the Excel template that contains smart markers and conditional formatting rules
        Workbook workbook = new Workbook("Template.xlsx");

        // Prepare the data source that will replace the smart markers
        // Example: a DataTable with sample data
        DataTable dt = new DataTable("Employees");
        dt.Columns.Add("Name", typeof(string));
        dt.Columns.Add("Department", typeof(string));
        dt.Columns.Add("Salary", typeof(double));

        dt.Rows.Add("John Doe", "Sales", 55000);
        dt.Rows.Add("Jane Smith", "Marketing", 62000);
        dt.Rows.Add("Bob Johnson", "IT", 72000);
        dt.Rows.Add("Alice Brown", "HR", 48000);

        // Create a WorkbookDesigner to process smart markers
        WorkbookDesigner designer = new WorkbookDesigner(workbook);

        // Set the data source for the smart markers
        designer.SetDataSource(dt);

        // Process the smart markers – this will replace placeholders with actual data
        // Conditional formatting rules defined in the template remain intact
        designer.Process();

        // Save the resulting workbook
        workbook.Save("Output.xlsx");
    }
}
