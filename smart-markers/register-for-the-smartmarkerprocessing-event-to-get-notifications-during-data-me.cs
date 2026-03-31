using System;
using System.Data;
using Aspose.Cells;

public class SmartMarkerDemo
{
    public static void Run()
    {
        // Load the Excel template that contains smart markers
        WorkbookDesigner designer = new WorkbookDesigner();
        designer.Workbook = new Workbook("Template.xlsx");

        // Create a sample data source (DataTable) to be merged into the template
        DataTable dt = new DataTable("Employees");
        dt.Columns.Add("Name", typeof(string));
        dt.Columns.Add("Age", typeof(int));
        dt.Rows.Add("John", 30);
        dt.Rows.Add("Jane", 28);

        // Bind the data source to the designer
        designer.SetDataSource(dt);

        // Process the smart markers and populate the worksheet
        designer.Process(true); // true = preserve unrecognized smart markers

        // Save the populated workbook
        designer.Workbook.Save("Result.xlsx");
    }
}

public class Program
{
    public static void Main()
    {
        SmartMarkerDemo.Run();
    }
}