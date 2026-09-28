// Title: Add Excel cell comments with Aspose.Cells smart markers from a DataTable in C#
// AI Prompts: Generate C# code that creates an Excel workbook, places a smart marker in a cell, attaches a comment containing a smart marker, binds a DataTable with Name and Comment columns to WorkbookDesigner, processes all markers, and saves the file as XLSX. | Show how to use Aspose.Cells WorkbookDesigner to replace smart markers inside both cell values and cell comments using data from a DataTable.
// Common Searches: how to bind a DataTable to Aspose.Cells smart markers for cell comments in C# | Aspose.Cells example of using smart markers inside Excel comments | C# code to generate Excel comments from a database column with smart markers | process smart markers in comments with WorkbookDesigner Aspose.Cells | populate Excel comment field using Aspose.Cells smart marker and DataTable
// Tags: Aspose.Cells comment insertion | data source binding for Excel comments | C# Excel comment generation | handle workbook markers in comments | generate XLSX with bound data

using System;
using System.Data;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsSmartMarkerComments
{
    // The example creates a new workbook, adds a smart marker to cell A2 and a comment containing a smart marker, builds a DataTable with Name and Comment columns, binds the table to WorkbookDesigner, processes all smart markers (including those inside comments), and saves the result as an XLSX file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // Set up a simple template with smart markers
            // -------------------------------------------------
            // Header
            sheet.Cells["A1"].PutValue("Product");
            // Cell that will be filled with product name using a smart marker
            sheet.Cells["A2"].PutValue("&=Name");

            // Add a comment to the same cell (A2) that contains a smart marker for the comment text
            int commentIdx = sheet.Comments.Add("A2");
            Comment comment = sheet.Comments[commentIdx];
            // The smart marker & =Comment will be replaced by the data source field "Comment"
            comment.Note = "&=Comment";

            // -------------------------------------------------
            // Prepare the data source (DataTable with Name and Comment columns)
            // -------------------------------------------------
            DataTable dt = new DataTable("Products");
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("Comment", typeof(string));

            dt.Rows.Add("Apple", "Fresh and juicy fruit.");
            dt.Rows.Add("Banana", "Rich in potassium.");
            dt.Rows.Add("Cherry", "Small and sweet.");

            // -------------------------------------------------
            // Configure WorkbookDesigner with the data source and process smart markers
            // -------------------------------------------------
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource(dt);
            // Process all smart markers in the workbook (including those inside comments)
            designer.Process();

            // -------------------------------------------------
            // Save the resulting workbook
            // -------------------------------------------------
            workbook.Save("SmartMarkerCommentsOutput.xlsx");
        }
    }
}
