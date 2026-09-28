// Title: Apply a StyleFlag to modify only the QuotePrefix property of a cell using Aspose.Cells for .NET
// AI Prompts: Write C# code that uses Aspose.Cells StyleFlag to enable QuotePrefix on a specific cell while keeping all other formatting unchanged. | Show how to create a Style with QuotePrefix set to true and apply it with a StyleFlag so that only the QuotePrefix attribute is updated. | Provide a complete example that saves an Excel workbook after applying a QuotePrefix‑only style flag to a target cell.
// Common Searches: Aspose.Cells C# set QuotePrefix without affecting other cell styles | How to use StyleFlag to change only the QuotePrefix attribute in an existing worksheet | C# Aspose.Cells apply QuotePrefix style flag to a single cell | Preserve existing cell formatting while enabling QuotePrefix with Aspose.Cells .NET
// Tags: Aspose.Cells StyleFlag QuotePrefix | C# cell QuotePrefix style flag | preserve cell formatting Aspose.Cells | apply single style attribute .NET | Excel workbook QuotePrefix style

using Aspose.Cells;
using System;

// Creates a workbook, enables the QuotePrefix flag on cell A1 using a Style and a StyleFlag so that only the QuotePrefix attribute is changed, then saves the file as QuotePrefixOnly.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Access cell A1 and put a value that looks like a number
        Cell cell = worksheet.Cells["A1"];
        cell.PutValue("12345");

        // Create a style and enable the QuotePrefix property
        Style style = workbook.CreateStyle();
        style.QuotePrefix = true;

        // Create a StyleFlag and enable only the QuotePrefix flag
        StyleFlag flag = new StyleFlag();
        flag.QuotePrefix = true; // all other flags remain false

        // Apply the style to the cell using the flag – only QuotePrefix will be updated
        cell.SetStyle(style, flag);

        // Save the workbook
        workbook.Save("QuotePrefixOnly.xlsx");
    }
}
