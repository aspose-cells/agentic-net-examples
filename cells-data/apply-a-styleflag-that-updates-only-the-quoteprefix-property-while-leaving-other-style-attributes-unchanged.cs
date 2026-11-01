// Title: Update only the QuotePrefix property of a cell with a StyleFlag in Aspose.Cells for .NET
// AI Prompts: Apply a StyleFlag that sets QuotePrefix to true on a specific cell while leaving all other formatting unchanged in C#. | Toggle the QuotePrefix attribute of a cell without resetting its existing style using Aspose.Cells. | Use Aspose.Cells to modify only the QuotePrefix flag of a cell style programmatically.
// Common Searches: Aspose.Cells C# set QuotePrefix on a cell without changing other styles | How to use StyleFlag to change only the QuotePrefix property in a workbook | Preserve existing cell formatting while enabling QuotePrefix in Aspose.Cells | Apply selective style flag for QuotePrefix in .NET Excel generation | Update cell QuotePrefix flag without affecting font color Aspose.Cells
// Tags: Aspose.Cells StyleFlag QuotePrefix | C# update cell QuotePrefix without altering style | selective style flag application Aspose.Cells | preserve existing cell formatting Aspose.Cells | modify single style attribute .NET Excel

using Aspose.Cells;

// The example creates a workbook, writes text to cell A1, retrieves its current style, sets the QuotePrefix property to true, creates a StyleFlag with only the QuotePrefix flag enabled, applies the style using this flag to keep all other formatting intact, and saves the file as output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Get cell A1
        Cell cell = sheet.Cells["A1"];
        cell.PutValue("Sample text");

        // Retrieve the current style of the cell
        Style style = cell.GetStyle();

        // Set only the QuotePrefix property to true
        style.QuotePrefix = true;

        // Create a StyleFlag and enable only the QuotePrefix flag
        StyleFlag flag = new StyleFlag();
        flag.QuotePrefix = true;   // other flags remain false (default)

        // Apply the style to the cell using the flag
        cell.SetStyle(style, flag);

        // Save the workbook
        workbook.Save("output.xlsx");
    }
}
