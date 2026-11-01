// Title: How to set bold and italic text in an Excel cell using the HtmlString property with Aspose.Cells for .NET (C#)
// AI Prompts: Assign '<b><i>Bold and Italic Text</i></b>' to a worksheet cell via the HtmlString property in C# using Aspose.Cells and save the workbook as XLSX. | Create a new workbook, set cell A1 HtmlString to bold‑italic HTML markup, and export the file with Aspose.Cells for .NET. | Use Aspose.Cells to embed HTML formatting in a cell and generate an Excel file programmatically in C#.
// Common Searches: asp.net aspose.cells htmlstring bold italic cell | c# set html formatted text in Excel cell using Aspose.Cells | how to apply HTML styling to a specific cell with Aspose.Cells for .NET | save workbook after using HtmlString to format cell content in Aspose.Cells
// Tags: Aspose.Cells HtmlString formatting | C# set bold italic HTML in Excel cell | write HTML to Excel cell with Aspose.Cells | export workbook as XLSX using Aspose.Cells

using Aspose.Cells;

// Creates a workbook, sets cell A1's HtmlString to '<b><i>Bold and Italic Text</i></b>', and saves the file as output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        var workbook = new Workbook();

        // Get the first worksheet
        var worksheet = workbook.Worksheets[0];

        // Access cell A1
        var cell = worksheet.Cells["A1"];

        // Assign HTML markup to display bold and italic text
        cell.HtmlString = "<b><i>Bold and Italic Text</i></b>";

        // Save the workbook
        workbook.Save("output.xlsx");
    }
}
