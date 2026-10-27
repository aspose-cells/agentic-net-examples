// Title: Using the HTML property in Aspose.Cells smart markers to render bold text with C#
// Common Searches: C# Aspose.Cells smart marker HTML property to make text bold | How to apply HTML formatting to a smart marker field in Excel using Aspose.Cells | Set bold HTML content in a cell via smart markers with WorkbookDesigner | Aspose.Cells example for rendering <b> tags from a DataSet in smart markers | Insert HTML property smart marker in Aspose.Cells workbook programmatically
// Tags: Aspose.Cells HTML smart marker | C# set smart marker HTML property | WorkbookDesigner process HTML formatting | Export bold text Excel using smart markers | smart marker HTML property example | DataSet to smart marker HTML content

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerHtmlExample
{
    // The sample creates a workbook, places a smart marker that uses the HTML property syntax (&="HTML" &="Text"), provides a DataSet with a '<b>Bold Text</b>' value, processes the marker with WorkbookDesigner, and saves the file. The resulting cell displays the text in bold using HTML formatting.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Insert a smart marker that uses the HTML property.
                // Syntax: &="HTML" &="FieldName"
                sheet.Cells["A1"].PutValue("&=\"HTML\" &=\"Text\"");

                // Prepare the data source as a DataSet with a DataTable.
                DataTable table = new DataTable();
                table.Columns.Add("Text", typeof(string));
                table.Rows.Add("<b>Bold Text</b>");

                DataSet dataSource = new DataSet();
                dataSource.Tables.Add(table);

                // Process the smart markers using WorkbookDesigner
                WorkbookDesigner designer = new WorkbookDesigner(workbook);
                designer.SetDataSource(dataSource);
                designer.Process();

                // Define output file path
                string outputPath = "SmartMarkerHtmlOutput.xlsx";

                // Ensure the directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the resulting workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
