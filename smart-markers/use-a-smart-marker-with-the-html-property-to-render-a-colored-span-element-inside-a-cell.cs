// Title: Insert a red HTML <span> into an Excel cell using Aspose.Cells smart markers in C#
// AI Prompts: Write C# code that creates a workbook, sets a cell's HtmlString to a red <span> containing a smart marker, binds a DataTable as the data source, processes the smart markers with WorkbookDesigner, and saves the file as XLSX. | Show how to use Aspose.Cells WorkbookDesigner to replace a smart marker inside an HTML span while preserving the span's color styling. | Demonstrate binding a DataTable to a smart marker, applying HTML formatting to a cell, and exporting the result with Aspose.Cells .NET. | Provide a step‑by‑step example of rendering colored text via an HTML span inside a cell using Aspose.Cells smart markers.
// Common Searches: aspnet aspose.cells how to embed html span with smart marker in a cell | c# set cell HtmlString with colored text using smart markers | bind datatable to workbookdesigner and process html smart markers | render red text in Excel cell via smart marker Aspose.Cells .NET | save workbook with styled html content from smart markers c#
// Tags: Aspose.Cells smart marker HTML span | WorkbookDesigner process smart markers C# | set cell HtmlString Aspose.Cells | bind DataTable to WorkbookDesigner | export workbook to XLSX with styled HTML | colored text in Excel using Aspose.Cells

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerHtmlDemo
{
    // Creates a workbook, inserts a red-colored HTML <span> with a smart marker (&=Data.Name&) into cell A1 via the HtmlString property, binds a DataTable as the data source, processes the marker using WorkbookDesigner, and saves the result as an XLSX file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Define a smart marker inside an HTML span.
                // The smart marker &=Data.Name& will be replaced with the value from the data source.
                // The span is styled with a red color.
                sheet.Cells["A1"].HtmlString = "<span style='color:#FF0000'>&=Data.Name&</span>";

                // Prepare a simple data source (DataTable) with a column "Name"
                DataTable dt = new DataTable("Data");
                dt.Columns.Add("Name", typeof(string));
                dt.Rows.Add("Aspose");
                dt.Rows.Add("Cells");
                dt.Rows.Add("SmartMarker");

                // Process the smart markers using WorkbookDesigner (the correct API)
                WorkbookDesigner designer = new WorkbookDesigner(workbook);
                designer.SetDataSource(dt);
                designer.Process();

                // Define output file path
                string outputPath = "SmartMarkerHtmlDemo.xlsx";

                // Ensure the directory for the output file exists (if any)
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
