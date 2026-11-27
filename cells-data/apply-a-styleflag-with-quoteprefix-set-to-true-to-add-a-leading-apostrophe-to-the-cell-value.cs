// Title: Add a leading apostrophe to an Excel cell using Aspose.Cells StyleFlag with QuotePrefix in C#
// Common Searches: how to display a leading apostrophe in an Excel cell with Aspose.Cells C# | apply QuotePrefix to a single cell using StyleFlag in Aspose.Cells .NET | C# code to enable QuotePrefix only for cell A1 in Aspose.Cells workbook | add apostrophe prefix to cell value without changing the text in Aspose.Cells | use StyleFlag to set QuotePrefix property for a cell in Aspose.Cells
// Tags: Aspose.Cells StyleFlag QuotePrefix C# | add leading apostrophe to Excel cell Aspose.Cells | apply cell style flag only QuotePrefix Aspose | single cell formatting with QuotePrefix Aspose.Cells | C# workbook save with styled cell Aspose

using Aspose.Cells;

// Demonstrates creating a workbook, setting a cell value, creating a style with QuotePrefix enabled, using a StyleFlag to apply only the QuotePrefix property to the cell, and saving the file as StyledCell.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        var workbook = new Workbook();

        // Access the first worksheet
        var worksheet = workbook.Worksheets[0];

        // Get the target cell (e.g., A1)
        var cell = worksheet.Cells["A1"];

        // Set the cell value (the apostrophe will be added by the style)
        cell.PutValue("Sample Text");

        // Create a new style and enable QuotePrefix
        var style = workbook.CreateStyle();
        style.QuotePrefix = true; // This adds a leading apostrophe when displayed

        // Create a StyleFlag and specify that only QuotePrefix should be applied
        var flag = new StyleFlag();
        flag.QuotePrefix = true;

        // Apply the style to the cell using the flag
        cell.SetStyle(style, flag);

        // Save the workbook to a file
        workbook.Save("StyledCell.xlsx");
    }
}
