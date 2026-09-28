// Title: Add a leading apostrophe to an Excel cell using StyleFlag.QuotePrefix in Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a workbook, sets "Hello World" in A1, and applies a StyleFlag with QuotePrefix enabled to show the value with a leading apostrophe. | Show how to apply only the QuotePrefix attribute to a cell style without affecting other formatting in Aspose.Cells. | Provide a complete example that saves the workbook after applying the QuotePrefix style flag to a specific cell.
// Common Searches: Aspose.Cells C# apply QuotePrefix to a single cell | How to display a leading apostrophe in Excel using Aspose.Cells StyleFlag | Selective style application QuotePrefix Aspose.Cells .NET example | Prepend apostrophe to cell value without changing other cell formats in Aspose.Cells
// Tags: StyleFlag QuotePrefix Aspose.Cells | prepend apostrophe Excel cell C# | apply selective cell style Aspose.Cells | save workbook with leading apostrophe .NET | cell formatting without affecting other attributes Aspose

using Aspose.Cells;

// Demonstrates creating a workbook, setting a value in cell A1, and using a StyleFlag with QuotePrefix set to true to prepend an apostrophe to the cell content, then saving the file as output.xlsx.
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

        // Set the cell value
        cell.PutValue("Hello World");

        // Create a style and enable QuotePrefix to add a leading apostrophe
        Style style = workbook.CreateStyle();
        style.QuotePrefix = true;

        // Define which style attributes to apply
        StyleFlag flag = new StyleFlag();
        flag.QuotePrefix = true;

        // Apply the style to the cell
        cell.SetStyle(style, flag);

        // Save the workbook
        workbook.Save("output.xlsx");
    }
}
