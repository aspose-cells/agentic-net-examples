// Title: How to keep a leading apostrophe in an Excel cell while applying a style using Aspose.Cells for .NET
// AI Prompts: Generate C# code that updates the formatting of cell A1 with Aspose.Cells without removing its leading apostrophe. | Show how to use Aspose.Cells StyleFlag to preserve the QuotePrefix when setting a cell style in .NET. | Provide an example that applies a style to a worksheet cell while keeping any existing single‑quote prefix.
// Common Searches: Aspose.Cells keep single quote in cell after SetStyle | C# Aspose.Cells StyleFlag QuotePrefix false example | Prevent Aspose.Cells from stripping leading apostrophe when formatting cells | How to preserve Excel cell text prefix when applying style with Aspose.Cells | SetStyle with StyleFlag to retain quote prefix in .NET
// Tags: Aspose.Cells StyleFlag QuotePrefix usage | retain single‑quote prefix in Excel cell | C# SetStyle with StyleFlag | Excel cell style application without stripping apostrophe | apply formatting while keeping text prefix Aspose

using Aspose.Cells;

// Loads a workbook, reads the style of cell A1, creates a StyleFlag with QuotePrefix set to false to keep any existing apostrophe, applies the style using SetStyle with the flag, and saves the workbook.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Get the current style of the target cell (e.g., A1)
        Style currentStyle = sheet.Cells["A1"].GetStyle();

        // Create a StyleFlag and set QuotePrefix to false
        // This ensures the existing apostrophe prefix (if any) is preserved
        StyleFlag flag = new StyleFlag();
        flag.QuotePrefix = false;

        // Apply the style with the specified flag to the cell
        sheet.Cells["A1"].SetStyle(currentStyle, flag);

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}
