// Title: C# – Render a colored HTML <span> in an Excel cell using Aspose.Cells smart markers
// Description: Demonstrates how to assign an HTML string with a red <span> and a smart marker (&=Data.Name&) to cell A1, bind a DataTable via WorkbookDesigner.SetDataSource, process the markers, and save the workbook as an .xlsx file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | smart markers | HtmlString property | C# Excel export | HTML span in cell | colored text Excel | WorkbookDesigner SetDataSource | DataTable binding | Excel template styling | Aspose.Cells .NET example
// Common Searches: Aspose.Cells HTML smart marker example | how to use HtmlString with smart markers in C# | render colored text in Excel using Aspose.Cells | setdatasource DataTable Aspose.Cells smart marker | C# Aspose.Cells template with HTML formatting
// Developer Intent: Insert styled HTML (colored <span>) into an Excel cell through a smart marker while generating the file programmatically.
// Use Cases: Create a report where each name from a DataTable appears in red text by embedding the value in an HTML <span> via a smart marker. | Design Excel templates that inject HTML‑styled fragments (color, bold, italics) from multiple data columns without using native conditional formatting. | Generate dynamic, data‑driven cell formatting (e.g., color‑coded status values) by combining HTML markup with smart markers.
// AI Prompts: Show how to change the <span> color based on a second column (e.g., Status) in the DataTable. | Explain how to place several smart markers with different HTML styles in the same worksheet. | Provide code that applies the HTML smart marker to an entire column using a loop instead of a single cell.

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerHtmlDemo
{
    // Demonstrates how to assign an HTML string with a red <span> and a smart marker (&=Data.Name&) to cell A1, bind a DataTable via WorkbookDesigner.SetDataSource, process the markers, and save the workbook as an .xlsx file using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Define a smart marker that contains an HTML <span> with a color style.
                // The smart marker &=Data.Name& will be replaced by the value from the data source.
                sheet.Cells["A1"].HtmlString = "<span style='color:#FF0000'>&=Data.Name&</span>";

                // Prepare a simple data source (DataTable) with a column named "Name"
                DataTable data = new DataTable("Data");
                data.Columns.Add("Name", typeof(string));
                data.Rows.Add("John Doe");
                data.Rows.Add("Jane Smith");

                // Process the smart markers using WorkbookDesigner (correct overload)
                WorkbookDesigner designer = new WorkbookDesigner(workbook);
                // In this API version the first argument is the data source name, the second is the data table
                designer.SetDataSource("Data", data);
                designer.Process();

                // Define output file path
                string outputPath = "SmartMarkerHtmlDemo.xlsx";

                // Ensure the output directory exists (if a directory part is present)
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
