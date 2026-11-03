// Title: Create an internal hyperlink to the range A1:B10 in the same worksheet with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add hyperlink that points to a cell range within the same sheet | how to use '#A1:B10' syntax for internal hyperlinks in Aspose.Cells workbook | C# code to create a clickable link to a range A1:B10 in an Aspose.Cells worksheet | setting hyperlink style blue underline for internal range link using Aspose.Cells .NET | save workbook after adding internal range hyperlink with Aspose.Cells
// Tags: Aspose.Cells worksheet range hyperlink API | C# add hyperlink to cell range in workbook | Aspose.Cells hyperlink address syntax '#A1:B10' | apply blue underline style to Aspose.Cells hyperlink | create same-sheet hyperlink using Aspose.Cells

using System;
using Aspose.Cells;
using System.Drawing;

// The example demonstrates how to add an internal hyperlink that points to the range A1:B10 within the same worksheet using Aspose.Cells for .NET. It sets the display text, applies blue underlined styling, and saves the workbook as HyperlinkToRange.xlsx.
class HyperlinkExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the cell where the hyperlink will be placed (e.g., A1)
            int hyperlinkRow = 0;      // Row index (0‑based)
            int hyperlinkColumn = 0;   // Column index (0‑based)

            // Add a hyperlink that points to the range A1:B10 in the same worksheet
            // The address uses the '#A1:B10' syntax for internal references
            worksheet.Hyperlinks.Add(
                hyperlinkRow,          // First row of the hyperlink range
                hyperlinkColumn,       // First column of the hyperlink range
                1,                     // Number of rows the hyperlink occupies
                1,                     // Number of columns the hyperlink occupies
                "#A1:B10");            // Internal address

            // Set the display text for the hyperlink cell
            worksheet.Cells[hyperlinkRow, hyperlinkColumn].PutValue("Go to A1:B10");

            // Apply typical hyperlink styling (blue and underlined)
            Style style = workbook.CreateStyle();
            style.Font.Color = Color.Blue;
            style.Font.Underline = FontUnderlineType.Single; // Correct enum value
            worksheet.Cells[hyperlinkRow, hyperlinkColumn].SetStyle(style);

            // Save the workbook to a file
            workbook.Save("HyperlinkToRange.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
