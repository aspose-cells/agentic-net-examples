// Title: Load an Excel template and assign a DataTable as a custom data source for smart markers using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells load Excel template and set DataTable as smart marker data source | Assign a DataTable to WorkbookDesigner smart markers in .NET | Example of using SetDataSource with a DataTable for smart markers in C# | Process smart markers after binding a DataTable with Aspose.Cells | How to populate smart markers from a DataTable in an Excel template using Aspose.Cells
// Tags: WorkbookDesigner data source binding | Aspose.Cells smart markers DataTable integration | C# load Excel template for smart markers | process smart markers and save workbook | custom data source for smart markers

using System;
using System.Data;
using Aspose.Cells;

// The example loads an Excel file (Template.xlsx) into a WorkbookDesigner, creates a DataTable named 'Employees' with Name and Age columns, fills it with sample rows, assigns the DataTable as the smart marker data source, processes the smart markers, and saves the resulting workbook as Result.xlsx.
class Program
{
    static void Main()
    {
        // Load the Excel template that contains smart markers
        WorkbookDesigner designer = new WorkbookDesigner();
        designer.Workbook = new Workbook("Template.xlsx");

        // Create a DataTable that will serve as the custom data source
        DataTable dt = new DataTable("Employees");
        dt.Columns.Add("Name", typeof(string));
        dt.Columns.Add("Age", typeof(int));

        // Populate the DataTable with sample data
        dt.Rows.Add("John", 30);
        dt.Rows.Add("Jane", 25);
        dt.Rows.Add("Bob", 40);

        // Assign the DataTable to the smart marker processor
        designer.SetDataSource(dt);

        // Process the smart markers in the template
        designer.Process();

        // Save the workbook with the populated data
        designer.Workbook.Save("Result.xlsx");
    }
}
