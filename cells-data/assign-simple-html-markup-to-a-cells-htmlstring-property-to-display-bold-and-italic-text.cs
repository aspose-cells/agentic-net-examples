// Title: Set bold and italic text in an Excel cell using the HtmlString property with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells how to write bold italic HTML into a worksheet cell | Assign HtmlString to Excel cell for rich text formatting using Aspose.Cells | Display bold and italic text in Excel file with Aspose.Cells .NET API | Use HtmlString property to format cell content as HTML in Aspose.Cells | Example of inserting <b><i> tags into an Excel cell with Aspose.Cells C#
// Tags: Aspose.Cells HtmlString formatting | C# set HTML content in Excel cell | Excel cell bold italic via HtmlString | Aspose.Cells write rich text to worksheet | Save workbook with HTML formatted cell

using Aspose.Cells;

// The example creates a new workbook, accesses cell A1 on the first worksheet, assigns the HTML string "<b><i>Bold and Italic Text</i></b>" to the cell's HtmlString property to render bold and italic text, and saves the workbook as output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Get the target cell (e.g., A1)
        Cell cell = sheet.Cells["A1"];

        // Assign HTML markup to display bold and italic text
        cell.HtmlString = "<b><i>Bold and Italic Text</i></b>";

        // Save the workbook to a file
        workbook.Save("output.xlsx");
    }
}
