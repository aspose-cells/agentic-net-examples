// Title: How to populate an Excel template containing smart markers from a DataTable using Aspose.Cells in C#
// Common Searches: C# Aspose.Cells load Excel template with smart markers and fill using DataTable | Using WorkbookDesigner to bind a DataTable to smart markers in an Excel file | Populate smart marker rows in an Excel workbook from a DataTable in .NET | Aspose.Cells process smart markers with DataTable source and save new workbook | Example code for filling smart markers from a DataTable in C#
// Tags: WorkbookDesigner DataTable binding | smart markers population using Aspose.Cells | Excel template processing with smart markers C# | Aspose.Cells save populated workbook | DataTable as smart marker source

using System;
using System.Data;
using Aspose.Cells;

// The example loads an Excel workbook that contains smart markers, creates a DataTable with employee data, assigns the DataTable as the data source for a WorkbookDesigner, processes the smart markers to generate a populated sheet, and saves the result as a new Excel file.
class SmartMarkerExample
{
    static void Main()
    {
        // Load the Excel template that contains smart markers
        string templatePath = "TemplateWithSmartMarkers.xlsx";
        Workbook workbook = new Workbook(templatePath);

        // Create and fill a DataTable that will be used as the data source
        DataTable dt = new DataTable("Employees");
        dt.Columns.Add("Name", typeof(string));
        dt.Columns.Add("Age", typeof(int));
        dt.Columns.Add("Department", typeof(string));

        dt.Rows.Add("John Doe", 30, "Finance");
        dt.Rows.Add("Jane Smith", 28, "HR");
        dt.Rows.Add("Bob Johnson", 35, "IT");

        // Set up the designer and assign the DataTable as a data source
        WorkbookDesigner designer = new WorkbookDesigner(workbook);
        designer.SetDataSource(dt);

        // Process the smart markers in the template
        designer.Process();

        // Save the populated workbook
        string outputPath = "PopulatedFromSmartMarkers.xlsx";
        workbook.Save(outputPath);
    }
}
