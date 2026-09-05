// Title: How to format a worksheet cell as a mixed fraction using a custom number format with Aspose.Cells for .NET (C#)
// AI Prompts: Write a C# snippet that uses Aspose.Cells to place 2.75 in cell A1 and apply a custom format that renders the value as a mixed number (e.g., 2 ¾). | Demonstrate setting a custom fraction pattern on a cell style with Aspose.Cells, then save the workbook as an .xlsx file. | Create an Aspose.Cells example that shows how to convert any decimal to a whole‑plus‑fraction representation by applying a custom number format to a worksheet cell.
// Common Searches: Aspose.Cells C# apply custom fraction format to a worksheet cell | Show integer part and fraction together in Excel using Aspose.Cells API | C# custom number format for displaying fractions as whole‑plus‑fraction in an .xlsx file
// Tags: Aspose.Cells custom mixed‑fraction number format | C# apply custom pattern to Excel cell using Aspose.Cells | save workbook with mixed number display Aspose.Cells | set cell style with custom fraction format in Aspose.Cells | convert decimal to whole‑plus‑fraction in .NET Excel export

using Aspose.Cells;

// Demonstrates creating a workbook, inserting the decimal 2.75 into cell A1, applying a custom number format that shows the value as a mixed fraction, and saving the result as MixedFraction.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule)
        var workbook = new Workbook();

        // Access the first worksheet
        var sheet = workbook.Worksheets[0];

        // Insert a decimal value that will be shown as a mixed fraction
        var cell = sheet.Cells["A1"];
        cell.PutValue(2.75); // Represents 2 ¾

        // Retrieve the cell's style
        var style = cell.GetStyle();

        // Apply a custom number format that displays fractions as mixed numbers
        // "# ??/??" shows an integer part followed by a fraction with up to two‑digit denominator
        style.Custom = "# ??/??";

        // Assign the modified style back to the cell
        cell.SetStyle(style);

        // Save the workbook (lifecycle rule)
        workbook.Save("MixedFraction.xlsx");
    }
}
