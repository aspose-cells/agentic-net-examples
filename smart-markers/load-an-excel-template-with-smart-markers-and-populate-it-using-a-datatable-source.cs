// Title: Populate an Excel Template with Smart Markers from a DataTable in C# (Aspose.Cells)
// Description: Learn how to load an Excel workbook that contains smart markers, bind a DataTable as the data source with WorkbookDesigner, process the markers, and save the populated file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | smart markers | C# | .NET | DataTable | WorkbookDesigner | Excel template | populate Excel | load workbook | process smart markers
// Common Searches: Aspose.Cells load Excel template with smart markers C# | WorkbookDesigner set DataTable source example | populate smart markers from DataTable Aspose.Cells | C# smart markers tutorial Aspose.Cells
// Developer Intent: Load an Excel template that contains smart markers and fill it using a DataTable as the data source.
// Use Cases: Generate an employee directory by mapping a DataTable to smart markers in a pre‑designed workbook. | Create invoices where each row of a DataTable populates customer and item smart markers. | Produce department summary sheets by binding multiple DataTables to a single smart‑marker workbook.
// AI Prompts: Write C# code that loads a smart‑marker Excel template, binds a DataSet with several DataTables to WorkbookDesigner, processes the markers, and saves the output file. | Explain how to bind a DataTable to Aspose.Cells WorkbookDesigner, apply custom number formats to smart markers, and skip rows when the data source is empty.

using System;
using System.Data;
using Aspose.Cells;

namespace SmartMarkerExample
{
    // Learn how to load an Excel workbook that contains smart markers, bind a DataTable as the data source with WorkbookDesigner, process the markers, and save the populated file using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Load the Excel template that contains smart markers
            Workbook workbook = new Workbook("TemplateWithSmartMarkers.xlsx");

            // Create and populate a DataTable that will serve as the data source
            DataTable dataTable = new DataTable("Employees");
            dataTable.Columns.Add("Name", typeof(string));
            dataTable.Columns.Add("Age", typeof(int));
            dataTable.Columns.Add("Department", typeof(string));

            dataTable.Rows.Add("John Doe", 30, "Sales");
            dataTable.Rows.Add("Jane Smith", 28, "Marketing");
            dataTable.Rows.Add("Mike Johnson", 35, "Engineering");

            // Initialize WorkbookDesigner with the loaded workbook
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // Set the DataTable as the data source for smart markers
            designer.SetDataSource(dataTable);

            // Process the smart markers and populate the worksheet with data
            designer.Process();

            // Save the populated workbook to a new file
            workbook.Save("PopulatedOutput.xlsx");
        }
    }
}
