// Title: How to render a red-colored HTML span in an Excel cell using Aspose.Cells smart markers (C#)
// Common Searches: Aspose.Cells C# render HTML string from smart marker in worksheet cell | display colored span HTML inside Excel cell using smart markers Aspose.Cells | set cell HtmlString after processing smart markers with WorkbookDesigner | use DataSet as data source for smart marker that contains HTML in Aspose.Cells | convert smart marker &=Data.Html to formatted HTML in Excel using C#
// Tags: smart-marker html rendering Aspose.Cells | WorkbookDesigner process html smart marker | cell HtmlString property C# | Excel cell colored span using Aspose.Cells | DataSet source for smart markers Aspose.Cells

using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, inserts a smart marker '&=Data.Html' into cell A1, provides a DataSet with an HTML string containing a red <span>, processes the marker with WorkbookDesigner, assigns the resulting HTML to the cell's HtmlString property to render the colored text, and saves the file as SmartMarkerHtmlOutput.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Insert a smart marker that will be replaced by the "Html" field value from the "Data" table
            sheet.Cells["A1"].PutValue("&=Data.Html");

            // Prepare data source as a DataSet with a single DataTable
            DataTable table = new DataTable("Data");
            table.Columns.Add("Html", typeof(string));
            table.Rows.Add("<span style='color:red'>Red Text</span>");

            DataSet dataSet = new DataSet();
            dataSet.Tables.Add(table);

            // Process smart markers using WorkbookDesigner
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource(dataSet);
            designer.Process();

            // Render the HTML string in the cell
            sheet.Cells["A1"].HtmlString = sheet.Cells["A1"].StringValue;

            // Save the workbook
            string outputPath = "SmartMarkerHtmlOutput.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
