// Title: How to assign a custom Name to a textbox shape in an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a workbook, adds a textbox shape, and sets its Name property to a custom identifier with Aspose.Cells. | Show how to rename an existing textbox shape in a worksheet and update its displayed text using Aspose.Cells for .NET. | Provide a complete example that positions a textbox over specific cells, assigns a unique name, and saves the workbook as an .xlsx file.
// Common Searches: Aspose.Cells C# set custom name for textbox shape in Excel | rename textbox shape programmatically in .NET workbook using Aspose.Cells | identify a specific textbox in an Excel file with Aspose.Cells API | C# example adding named textbox to worksheet with Aspose.Cells | assign unique identifier to Excel shape for later manipulation Aspose.Cells
// Tags: add textbox shape Aspose.Cells C# | set shape Name property Aspose.Cells | named textbox Excel Aspose.Cells | shape identification Aspose.Cells workbook | save workbook as xlsx Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a new workbook, adds a textbox shape positioned over cells B2:C4, assigns it a unique Name ("UniqueTextbox_001"), sets its displayed text, and saves the file as Output.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a textbox shape to the worksheet.
            // Parameters: upper left row, upper left column, top offset, left offset, width, height (in pixels)
            // This places the textbox roughly over cells B2:C4.
            Shape textbox = sheet.Shapes.AddTextBox(
                1,    // upper left row (B2, zero‑based)
                1,    // upper left column (B2)
                0,    // top offset (pixels)
                0,    // left offset (pixels)
                200,  // width (pixels)
                100   // height (pixels)
            );

            // Assign a unique name for later identification and manipulation
            textbox.Name = "UniqueTextbox_001";

            // Set the displayed text
            textbox.Text = "Sample Text";

            // Save the workbook (lifecycle rule: save)
            workbook.Save("Output.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
