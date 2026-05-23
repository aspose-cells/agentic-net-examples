using System;
using System.Data;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class EmbedCommentsWithSmartMarkers
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a sample value to the cell (optional, just to have visible content)
        sheet.Cells["A1"].PutValue("Product A");

        // Add a comment to cell A1
        int commentIdx = sheet.Comments.Add("A1");
        Comment comment = sheet.Comments[commentIdx];

        // Use a smart marker in the comment text; it will be replaced by the data source field "CommentText"
        comment.Note = "&=$CommentText";

        // Prepare a data source (DataTable) with a column that matches the smart marker name
        DataTable dt = new DataTable("Data");
        dt.Columns.Add("CommentText", typeof(string));
        dt.Rows.Add("This product is top‑seller for Q1.");

        // Set up WorkbookDesigner, assign the data source, and process smart markers
        WorkbookDesigner designer = new WorkbookDesigner(workbook);
        designer.SetDataSource(dt);
        designer.Process(); // replaces smart markers with actual data

        // Save the resulting workbook
        workbook.Save("CommentsWithSmartMarkers.xlsx");
    }
}