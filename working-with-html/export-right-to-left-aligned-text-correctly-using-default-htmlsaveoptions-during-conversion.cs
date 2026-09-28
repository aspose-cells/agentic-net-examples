// Title: Export right-to-left Hebrew text to HTML with Aspose.Cells default HtmlSaveOptions (C#)
// AI Prompts: Write C# code that inserts Hebrew text into a worksheet cell, applies a RightToLeft text direction and right alignment, and saves the workbook as HTML using the default HtmlSaveOptions. | Show how to keep right-to-left cell formatting when converting an Excel file to HTML with Aspose.Cells without customizing save options. | Create a minimal Aspose.Cells example that demonstrates RTL text direction in the generated HTML output.
// Common Searches: Aspose.Cells export RTL text to HTML using default HtmlSaveOptions in C# | How to preserve Hebrew right-to-left alignment when saving Excel as HTML with Aspose.Cells | C# convert Excel workbook with Arabic RTL cells to HTML preserving direction | Default HtmlSaveOptions keep text direction Aspose.Cells example
// Tags: RTL text export to HTML Aspose.Cells C# | HtmlSaveOptions default preserve text direction | TextDirectionType.RightToLeft cell style Aspose.Cells | right-to-left alignment Excel to HTML conversion | Hebrew cell style Aspose.Cells example

using Aspose.Cells;
using System;

// Creates a workbook, writes Hebrew text into cell A1, sets the cell's TextDirection to RightToLeft and HorizontalAlignment to Right, and saves the workbook as an HTML file using Aspose.Cells default HtmlSaveOptions, preserving the right-to-left alignment in the output.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Insert right-to-left text (e.g., Hebrew) into a cell
            Cell cell = sheet.Cells["A1"];
            cell.PutValue("שלום"); // "Hello" in Hebrew

            // Configure the cell style for RTL alignment
            Style style = cell.GetStyle();
            style.HorizontalAlignment = TextAlignmentType.Right; // Align text to the right
            style.TextDirection = TextDirectionType.RightToLeft; // Set cell direction to RTL
            cell.SetStyle(style);

            // Use default HtmlSaveOptions (RTL direction is handled by cell style)
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();

            // Save the workbook as an HTML file
            workbook.Save("RtlOutput.html", saveOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
