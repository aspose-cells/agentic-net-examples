// Title: How to bind a custom DataTable to WorkbookDesigner for smart marker processing in Aspose.Cells for .NET
// Common Searches: Aspose.Cells set DataTable as data source for WorkbookDesigner in C# | C# example of using WorkbookDesigner with a custom DataTable for smart markers | process smart markers in an Excel template using a DataTable with Aspose.Cells | assign a DataTable to WorkbookDesigner before calling Process in Aspose.Cells | load Excel template and bind DataTable to WorkbookDesigner for smart marker replacement
// Tags: WorkbookDesigner set DataTable source | Aspose.Cells smart markers DataTable binding | C# bind DataTable to Excel template | process smart markers using WorkbookDesigner | custom data source for Excel smart markers

using System;
using System.Data;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a DataTable with employee data, loads an Excel template containing smart markers, assigns the DataTable as the data source for WorkbookDesigner, processes the smart markers, and saves the result as output.xlsx.
class Program
{
    static void Main()
    {
        // Create a custom DataTable and populate it with sample data
        DataTable dt = new DataTable("Employees");
        dt.Columns.Add("Name", typeof(string));
        dt.Columns.Add("Age", typeof(int));
        dt.Columns.Add("Department", typeof(string));

        dt.Rows.Add("Alice", 30, "HR");
        dt.Rows.Add("Bob", 45, "IT");
        dt.Rows.Add("Charlie", 28, "Finance");

        // Load the Excel template that contains smart markers
        WorkbookDesigner designer = new WorkbookDesigner();
        designer.Workbook = new Workbook("template.xlsx"); // ensure template.xlsx exists

        // Assign the custom DataTable as the data source for the designer
        designer.SetDataSource(dt);

        // Process the smart markers in the template using the provided data source
        designer.Process();

        // Save the resulting workbook
        designer.Workbook.Save("output.xlsx");
    }
}
