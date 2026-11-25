// Title: Render a colored <span> inside an Excel cell using a smart marker with the Html property in Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# smart marker html property to display colored text in a cell | how to insert an HTML span with red color into Excel using WorkbookDesigner | render HTML markup from a DataSet into an Excel cell via Aspose.Cells smart markers | C# example of using &=Html smart marker in an Aspose.Cells template workbook | populate an Excel template with HTML string using Aspose.Cells smart markers
// Tags: smart marker html property Aspose.Cells | render html in excel cell c# | workbookdesigner dataset data source | colored span markup excel Aspose | excel template html rendering c#

using System;
using System.Collections.Generic;
using System.Data;
using Aspose.Cells;

// The example creates a workbook template, places the smart marker "&=Html" in cell A1, supplies a DataSet with an Html column containing a red-colored <span> HTML string, assigns the data source to WorkbookDesigner, processes the smart markers so the HTML is rendered inside the cell, and saves the result as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook that will serve as the template
            Workbook templateWorkbook = new Workbook();
            Worksheet sheet = templateWorkbook.Worksheets[0];

            // Insert a smart marker with the HTML property into cell A1
            // The marker "&=Html" tells Aspose.Cells to replace it with the value of the "Html" property from the data source
            sheet.Cells["A1"].PutValue("&=Html");

            // Initialize WorkbookDesigner which processes smart markers
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = templateWorkbook
            };

            // Prepare the data source: a DataSet containing a DataTable with an "Html" column
            DataSet dataSet = new DataSet();
            DataTable table = new DataTable("HtmlData");
            table.Columns.Add("Html", typeof(string));
            table.Rows.Add("<span style='color:red'>Colored Text</span>");
            dataSet.Tables.Add(table);

            // Assign the data source to the designer
            designer.SetDataSource(dataSet);

            // Process the smart markers – the HTML string will be rendered inside the cell
            designer.Process();

            // Save the resulting workbook
            string outputPath = "output.xlsx";
            designer.Workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

// Data class (kept for reference; not used directly in this example)
public class HtmlData
{
    public string? Html { get; set; }
}
