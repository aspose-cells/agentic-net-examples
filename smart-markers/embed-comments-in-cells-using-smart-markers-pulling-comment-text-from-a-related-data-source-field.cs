// Title: Add Excel cell comments with smart markers from a DataTable using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that attaches a comment containing a smart marker to a worksheet cell and fills the comment text from a DataTable column with Aspose.Cells. | Demonstrate how to invoke WorkbookDesigner to process smart markers in both cell values and cell comments based on a DataTable source. | Explain the steps to create a DataTable, bind it to WorkbookDesigner, and generate an Excel file where comments are populated dynamically via smart markers.
// Common Searches: asp.net how to insert a comment with a smart marker using Aspose.Cells C# example | populate Excel comment text from DataTable column with Aspose.Cells WorkbookDesigner | smart marker for cell comment Aspose.Cells .NET tutorial | generate Excel file with dynamic comments from database using Aspose.Cells smart markers
// Tags: smart-marker comment insertion Aspose.Cells C# | WorkbookDesigner populate comments from DataTable | dynamic Excel comment generation Aspose.Cells | cell comment smart marker processing .NET

using System;
using System.Data;
using Aspose.Cells;

// The example creates a DataTable with Name and Comment columns, places a smart marker for Name in cell A2, adds a comment to the same cell containing a smart marker for Comment, processes all markers with WorkbookDesigner, and saves the workbook as EmployeesWithComments.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Prepare the data source with a field for the comment text.
            DataTable dt = new DataTable("Employees");
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("Comment", typeof(string));
            dt.Rows.Add("John Doe", "Top performer");
            dt.Rows.Add("Jane Smith", "Needs improvement");
            dt.Rows.Add("Bob Johnson", "Excellent attendance");

            // Create a new workbook and obtain the first worksheet.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Employees";

            // Add a header for the Name column.
            sheet.Cells["A1"].PutValue("Name");

            // Insert a smart marker for the Name field. This will be expanded for each data row.
            sheet.Cells["A2"].PutValue("&=Name");

            // Add a comment to the same cell (A2) that contains a smart marker for the Comment field.
            // When the smart markers are processed, the comment text will be replaced with the value from the data source.
            int commentIndex = sheet.Comments.Add("A2");
            Comment comment = sheet.Comments[commentIndex];
            comment.Note = "&=Comment";

            // Use WorkbookDesigner to process the smart markers with the provided data source.
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource(dt);
            designer.Process();

            // Save the resulting workbook.
            workbook.Save("EmployeesWithComments.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
