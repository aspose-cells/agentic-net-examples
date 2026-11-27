// Title: How to use Aspose.Cells StyleFlag to change only the QuotePrefix of a cell in C#
// Common Searches: Aspose.Cells C# set QuotePrefix without affecting other cell styles | Use StyleFlag to update only QuotePrefix property in Aspose.Cells workbook | Preserve existing cell formatting while enabling QuotePrefix in Aspose.Cells | C# Aspose.Cells SetStyle with selective StyleFlag for QuotePrefix | Apply QuotePrefix to a cell using Aspose.Cells StyleFlag example
// Tags: StyleFlag selective update Aspose.Cells | QuotePrefix cell style Aspose.Cells | preserve existing formatting Aspose.Cells | SetStyle with flag Aspose.Cells | Excel cell style modification Aspose.Cells

using System.Drawing;
using Aspose.Cells;

// Creates a workbook, applies a red font to cell A1, then defines a StyleFlag that enables only the QuotePrefix flag. Using SetStyle with this flag updates just the QuotePrefix property while all other cell formatting remains unchanged, and saves the workbook as Output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Access the target cell (e.g., A1)
        Cell cell = sheet.Cells["A1"];

        // Example: set an initial style attribute to show it remains unchanged
        Style initialStyle = cell.GetStyle();
        initialStyle.Font.Color = Color.Red;
        cell.SetStyle(initialStyle);

        // Prepare a new style with only the QuotePrefix property set
        Style quoteStyle = cell.GetStyle();
        quoteStyle.QuotePrefix = true;

        // Define a StyleFlag that enables updating only QuotePrefix
        StyleFlag flag = new StyleFlag();
        flag.QuotePrefix = true; // other flags remain false

        // Apply the style using the flag; only QuotePrefix will be updated
        cell.SetStyle(quoteStyle, flag);

        // Save the workbook
        workbook.Save("Output.xlsx");
    }
}
