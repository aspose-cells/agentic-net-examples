using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerComments
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Put a sample value in cell A1 (the cell that will have a comment)
            sheet.Cells["A1"].PutValue("Product A");

            // Add a comment to cell A1
            int commentIndex = sheet.Comments.Add("A1");
            Comment comment = sheet.Comments[commentIndex];

            // Set the comment text to a smart marker that will be replaced by data source field "CommentText"
            comment.Note = "&=$CommentText";

            // Prepare a data source (DataTable) with a column that matches the smart marker name
            DataTable dt = new DataTable("Data");
            dt.Columns.Add("CommentText", typeof(string));
            dt.Rows.Add("This is a dynamically generated comment for Product A.");

            // Initialize WorkbookDesigner, assign the workbook and set the data source
            WorkbookDesigner designer = new WorkbookDesigner();
            designer.Workbook = workbook;
            designer.SetDataSource(dt);

            // Process the smart markers – this will replace the smart marker in the comment with the actual text
            designer.Process();

            // Save the resulting workbook
            workbook.Save("SmartMarkerCommentsOutput.xlsx");
        }
    }
}