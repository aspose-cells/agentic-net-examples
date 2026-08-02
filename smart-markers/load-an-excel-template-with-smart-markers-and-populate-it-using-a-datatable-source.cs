// Title: C# – Populate an Excel template with smart markers from a DataTable using Aspose.Cells
// Description: Loads a workbook that contains smart markers, creates a matching DataTable, assigns it as the data source to a WorkbookDesigner, processes the markers, and saves the filled workbook. Demonstrates end‑to‑end automation of Excel templates with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | smart markers | WorkbookDesigner | DataTable source | populate Excel template | process smart markers | Excel automation | template filling | generate report
// Common Searches: Aspose.Cells C# load template with smart markers | How to fill smart markers from a DataTable | WorkbookDesigner SetDataSource example | Populate Excel template using Aspose.Cells | Smart markers C# sample code
// Developer Intent: The developer needs to read an existing Excel file that contains smart markers, bind a DataTable to those markers, generate the final document, and save it programmatically.
// Use Cases: Create employee directories by mapping DataTable rows to smart markers in a pre‑designed sheet. | Generate batch invoices where each record populates customer and order details via smart markers. | Automate HR onboarding forms by filling template placeholders with data pulled from a database.
// AI Prompts: Show a C# snippet that reads a DataTable from SQL Server and uses WorkbookDesigner to replace smart markers in an Excel template. | Explain how to handle multiple DataTables with separate smart‑marker groups in the same workbook. | Provide code to set the output file name dynamically based on a column value from the DataTable when processing smart markers.

using System;
using System.Data;
using Aspose.Cells;

// Loads a workbook that contains smart markers, creates a matching DataTable, assigns it as the data source to a WorkbookDesigner, processes the markers, and saves the filled workbook. Demonstrates end‑to‑end automation of Excel templates with Aspose.Cells for .NET.
class SmartMarkerExample
{
    static void Main()
    {
        // Load the Excel template that contains smart markers (e.g., &=$Name, &=$Age, &=$Department)
        Workbook workbook = new Workbook("TemplateWithSmartMarkers.xlsx");

        // Create a DataTable that matches the smart marker fields
        DataTable dt = new DataTable("Employees");
        dt.Columns.Add("Name", typeof(string));
        dt.Columns.Add("Age", typeof(int));
        dt.Columns.Add("Department", typeof(string));

        // Populate the DataTable with sample data
        dt.Rows.Add("John Doe", 30, "Sales");
        dt.Rows.Add("Jane Smith", 28, "Marketing");
        dt.Rows.Add("Mike Johnson", 35, "Engineering");

        // Initialize WorkbookDesigner with the loaded workbook
        WorkbookDesigner designer = new WorkbookDesigner(workbook);

        // Set the DataTable as the data source for the smart markers
        designer.SetDataSource(dt);

        // Process the smart markers and fill the worksheet with data from the DataTable
        designer.Process();

        // Save the resulting workbook
        workbook.Save("PopulatedOutput.xlsx");
    }
}
