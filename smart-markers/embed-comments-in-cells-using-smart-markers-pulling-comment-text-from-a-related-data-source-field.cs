// Title: Create Excel cell comments from a DataTable using Aspose.Cells smart markers in C#
// Common Searches: how to generate Excel cell comments from a DataTable with Aspose.Cells C# | using smart markers to populate comments in Aspose.Cells workbook | Aspose.Cells C# add comment text via /*{Comment}*/ marker | populate Excel comments dynamically with WorkbookDesigner and data source
// Tags: smart markers add cell comments C# | WorkbookDesigner populate comments from DataTable | Aspose.Cells embed comment marker syntax | generate Excel comments using data source | C# create comments with smart markers

using System;
using System.Data;
using Aspose.Cells;

namespace SmartMarkerCommentsExample
{
    // The example creates a workbook, inserts a {Value} smart marker in cell A1, adds a comment containing the /*{Comment}*/ marker, builds a DataTable with Value and Comment columns, processes the markers with WorkbookDesigner, and saves the result as SmartMarkerCommentsOutput.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            // 1. Create a new workbook (template) and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // 2. Insert a smart marker for the cell value in A1
            //    The marker {Value} will be replaced with data from the "Value" column
            sheet.Cells["A1"].PutValue("{Value}");

            // 3. Add a comment to the same cell with a smart marker for the comment text
            //    The syntax /*{Comment}*/ tells Aspose.Cells to replace the comment with data from the "Comment" column
            Comment comment = sheet.Comments[sheet.Comments.Add("A1")];
            comment.Note = "/*{Comment}*/";

            // 4. Prepare the data source (DataTable) that contains the values and comments
            DataTable dt = new DataTable();
            dt.Columns.Add("Value", typeof(string));
            dt.Columns.Add("Comment", typeof(string));

            // Sample rows
            dt.Rows.Add("Apple", "This is a fruit.");
            dt.Rows.Add("Carrot", "This is a vegetable.");
            dt.Rows.Add("Blue", "This is a color.");

            // 5. Use WorkbookDesigner to process smart markers with the data source
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource(dt);
            designer.Process(); // This replaces the smart markers and creates the comments

            // 6. Save the resulting workbook
            workbook.Save("SmartMarkerCommentsOutput.xlsx");
        }
    }
}
