// Title: Render bold HTML text in an Excel cell using Aspose.Cells smart markers (C#)
// AI Prompts: Write C# code that sets a smart marker's HtmlString to "<b>&=FieldName</b>", binds a DataTable to WorkbookDesigner, processes the marker, and saves the workbook as an .xlsx file. | Show how to use Aspose.Cells WorkbookDesigner to apply HTML bold formatting to a smart marker and generate a spreadsheet with bold text.
// Common Searches: how to embed <b> tags in a smart marker for Aspose.Cells C# | Aspose.Cells smart marker HtmlString property usage example | process smart markers with HTML content to produce bold text in Excel using C# | C# code sample for binding DataTable to WorkbookDesigner and saving as .xlsx
// Tags: smart marker html bold Aspose.Cells | WorkbookDesigner set HtmlString C# | export bold text to xlsx via smart markers | bind DataTable to WorkbookDesigner Aspose.Cells | HTML rendering in Excel cells using Aspose.Cells

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

namespace SmartMarkerExample
{
    // The example creates a workbook, assigns "<b>&=Name</b>" to cell A1 via the HtmlString property, supplies a DataTable with a Name column, processes the smart marker with WorkbookDesigner, and saves the result as SmartMarkerBoldOutput.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Define a smart marker with HTML bold tags in cell A1
                sheet.Cells["A1"].HtmlString = "<b>&=Name</b>";

                // Prepare sample data source
                DataTable dt = new DataTable();
                dt.Columns.Add("Name", typeof(string));
                dt.Rows.Add("Bold Text Example");

                // Process smart markers
                WorkbookDesigner designer = new WorkbookDesigner(workbook);
                designer.SetDataSource(dt);
                designer.Process();

                // Ensure output directory exists
                string outputPath = "SmartMarkerBoldOutput.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the resulting workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
