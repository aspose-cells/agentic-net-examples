// Title: Create an internal hyperlink that jumps to a cell range (A1:B10) within the same worksheet using Aspose.Cells for .NET
// AI Prompts: Place a link in cell A1 that jumps to the range A1:B10 on the same sheet using Aspose.Cells. | Assign a custom label to the link and write the workbook to HyperlinkToRange.xlsx. | Redirect the link to a different range such as C5:D15 while preserving the # syntax.
// Common Searches: Aspose.Cells create link that jumps to A1:B10 in the same sheet | C# use # syntax for worksheet‑internal hyperlink with Aspose.Cells | set visible text for a worksheet link in Aspose.Cells | save workbook after adding a range link using Aspose.Cells
// Tags: Aspose.Cells add sheet‑internal link | Aspose.Cells link to specific cell range | Aspose.Cells customize link display text | Aspose.Cells save workbook with link

using System;
using Aspose.Cells;

// Shows how to add a sheet‑internal hyperlink from cell A1 to the range A1:B10, set custom display text, and save the workbook using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a hyperlink to cell A1 that points to the range A1:B10 in the same sheet
        // The address uses the "#" syntax for internal links
        worksheet.Hyperlinks.Add("A1", 1, 1, "#A1:B10");

        // Set the text that will be displayed for the hyperlink
        worksheet.Hyperlinks[0].TextToDisplay = "Go to A1:B10";

        // Save the workbook to a file
        workbook.Save("HyperlinkToRange.xlsx");
    }
}
