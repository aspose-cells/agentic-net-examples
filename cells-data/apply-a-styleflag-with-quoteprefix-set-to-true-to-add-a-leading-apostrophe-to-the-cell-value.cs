// Title: Add a leading apostrophe to an Excel cell using Aspose.Cells Style.QuotePrefix in C#
// Common Searches: how to force Excel cell to treat text as literal using Aspose.Cells C# | apply QuotePrefix on a cell style with Aspose.Cells .NET API | example of adding a leading apostrophe to a cell in Aspose.Cells | C# code to set cell style flag for text prefix in Excel workbook | Aspose.Cells method to display leading single quote in cell
// Tags: Aspose.Cells QuotePrefix property usage | C# add leading single quote to cell value | Excel cell style flag for literal text | saving workbook with quoted cell in Aspose.Cells | cell formatting to force text interpretation in .NET

using System;
using Aspose.Cells;

// The example creates a new workbook, writes "Hello" to cell A1, enables the QuotePrefix style flag to prepend an apostrophe to the cell value, and saves the workbook as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle create rule)
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Access cell A1 and set its value
            Cell cell = sheet.Cells["A1"];
            cell.PutValue("Hello");

            // Apply QuotePrefix flag to add a leading apostrophe via the cell's style
            Style style = cell.GetStyle();
            // Correct property name for the quote prefix flag
            style.QuotePrefix = true;
            cell.SetStyle(style);

            // Save the workbook (lifecycle save rule)
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
