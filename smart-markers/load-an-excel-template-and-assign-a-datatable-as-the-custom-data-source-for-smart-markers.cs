// Title: Load an Excel template and bind a DataTable as a custom data source for smart markers with Aspose.Cells for .NET (C#)
// Common Searches: how to use WorkbookDesigner to set a DataTable as smart marker source in C# | assigning a DataTable to smart markers in an Excel template using Aspose.Cells .NET | processing smart markers with a custom DataTable source in Aspose.Cells C# example | load template.xlsx and apply smart markers from a DataTable with Aspose.Cells | Aspose.Cells smart markers custom data source from DataTable tutorial
// Tags: WorkbookDesigner SetDataSource with DataTable | smart markers custom data source in .NET | load Excel template for smart marker processing | process smart markers using Aspose.Cells C# | bind DataTable to smart markers Aspose.Cells

using Aspose.Cells;
using System.Data;

// The example loads an Excel file (Template.xlsx), creates a DataTable named Employees with Name and Age columns, populates it, assigns the DataTable to a WorkbookDesigner as a custom data source for smart markers, processes the markers, and saves the result to Result.xlsx.
class Program
{
    static void Main()
    {
        // Load the Excel template
        Workbook workbook = new Workbook("Template.xlsx");

        // Create and populate a DataTable (example data)
        DataTable dt = new DataTable("Employees");
        dt.Columns.Add("Name", typeof(string));
        dt.Columns.Add("Age", typeof(int));
        dt.Rows.Add("John Doe", 30);
        dt.Rows.Add("Jane Smith", 28);

        // Assign the DataTable as a custom data source for smart markers
        WorkbookDesigner designer = new WorkbookDesigner(workbook);
        designer.SetDataSource(dt);               // or designer.SetDataSource("Employees", dt);
        designer.Process();                       // Apply smart markers

        // Save the processed workbook
        workbook.Save("Result.xlsx");
    }
}
