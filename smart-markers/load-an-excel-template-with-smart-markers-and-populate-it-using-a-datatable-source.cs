// Title: Load an Excel template with smart markers and fill it using a DataTable in C# with Aspose.Cells
// AI Prompts: Generate C# code that opens a .xlsx template containing Aspose.Cells smart markers, creates a DataTable with product data, binds it to WorkbookDesigner, processes the markers, and saves the result. | Show how to use WorkbookDesigner.SetDataSource to connect a DataTable to smart markers and produce a populated workbook in a .NET application. | Provide a step‑by‑step example that demonstrates loading a smart‑marker template, populating it from a DataTable, and writing the output file with Aspose.Cells.
// Common Searches: Aspose.Cells C# load Excel template with smart markers and populate from DataTable | How to bind a DataTable to WorkbookDesigner for smart marker processing in .NET | C# example of filling smart markers in an .xlsx file using Aspose.Cells WorkbookDesigner | Populate Excel smart markers using a DataTable source with Aspose.Cells | Save populated workbook after processing smart markers in C# Aspose.Cells
// Tags: Aspose.Cells WorkbookDesigner DataTable binding | populate smart markers from DataTable | C# Excel template smart marker processing | load Excel template with smart markers Aspose.Cells | save populated workbook Aspose.Cells C#

using System;
using System.Data;
using Aspose.Cells;

namespace SmartMarkerExample
{
    // // Loads an Excel template containing Aspose.Cells smart markers, creates a DataTable with product information, binds it to WorkbookDesigner, processes the markers, and saves the populated workbook as PopulatedResult.xlsx.
    class Program
    {
        static void Main()
        {
            // Load the Excel template that contains smart markers (e.g., &=$Data.Name)
            Workbook workbook = new Workbook("TemplateWithSmartMarkers.xlsx");

            // Initialize the WorkbookDesigner with the loaded workbook
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // Create and populate a DataTable that will serve as the data source
            DataTable dataTable = new DataTable("Data");
            dataTable.Columns.Add("Name", typeof(string));
            dataTable.Columns.Add("Quantity", typeof(int));
            dataTable.Columns.Add("Price", typeof(decimal));

            dataTable.Rows.Add("Apple", 10, 0.55m);
            dataTable.Rows.Add("Banana", 20, 0.30m);
            dataTable.Rows.Add("Cherry", 15, 1.20m);

            // Bind the DataTable to the designer
            designer.SetDataSource(dataTable);

            // Process the smart markers and populate the worksheet with data
            designer.Process();

            // Save the populated workbook to a new file
            workbook.Save("PopulatedResult.xlsx");
        }
    }
}
