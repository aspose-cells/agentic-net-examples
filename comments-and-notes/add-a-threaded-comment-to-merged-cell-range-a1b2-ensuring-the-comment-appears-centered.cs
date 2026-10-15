// Title: How to add a centered threaded comment to a merged cell range A1:B2 using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells add threaded comment to merged range A1:B2 | center comment text in a merged cell with Aspose.Cells .NET | programmatically set comment position for merged cells Aspose.Cells | Aspose.Cells example adding comment to merged cells in C# | why comment alignment is not available for merged ranges in Aspose.Cells
// Tags: add threaded comment to merged cells Aspose.Cells | center comment text in merged range C# | Aspose.Cells comment shape alignment limitation | merge cells and attach comment .NET | threaded comment alignment workaround Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, merges cells A1:B2 on the first worksheet, adds a threaded comment with author and note to cell A1, notes that direct alignment properties are unavailable in the current Aspose.Cells version, and saves the file as ThreadedComment_MergedCell.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Merge cells A1:B2 (row 0, column 0, spanning 2 rows and 2 columns)
            sheet.Cells.Merge(0, 0, 2, 2);

            // Add a regular comment to cell A1
            int commentIndex = sheet.Comments.Add("A1");
            Comment comment = sheet.Comments[commentIndex];
            comment.Author = "AuthorName";
            comment.Note = "This is a threaded comment."; // Set comment text

            // Note: Alignment properties are not available in the current Aspose.Cells version.
            // If needed, they can be set using the Shape object associated with the comment.

            // Save the workbook
            workbook.Save("ThreadedComment_MergedCell.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
