// Title: Apply left horizontal alignment to textbox shape text using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a textbox on a worksheet with Aspose.Cells and sets its TextHorizontalAlignment property to Left. | Show how to modify an existing Aspose.Cells workbook to change the horizontal alignment of a textbox's text to left. | Provide a step‑by‑step example that adds a textbox shape, assigns text, and left‑aligns the content using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# left align text inside a textbox shape in Excel | How to set TextHorizontalAlignment to Left for a textbox using Aspose.Cells .NET | C# code to add a textbox to an Excel worksheet and align its text to the left with Aspose.Cells | Changing horizontal alignment of textbox content in an existing workbook with Aspose.Cells for .NET
// Tags: Aspose.Cells textbox TextHorizontalAlignment left | C# set textbox horizontal alignment Aspose.Cells | Excel shape text alignment Aspose.Cells .NET | Add textbox shape left align text Aspose.Cells | Modify textbox alignment in existing workbook Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// // Loads or creates an Excel workbook, adds a textbox shape to the first worksheet, sets its text, applies left horizontal alignment via TextHorizontalAlignment, and saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Load existing workbook or create a new one if the file is missing.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                workbook.Worksheets.Add("Sheet1");
            }

            // Work with the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Add a textbox shape (creates a new one each run).
            // Parameters: upper left row, upper left column, top, left, height, width
            Shape textbox = sheet.Shapes.AddTextBox(0, 0, 0, 0, 100, 200);
            textbox.Text = "Sample Text";

            // Apply left alignment to the entire text within the textbox.
            textbox.TextHorizontalAlignment = TextAlignmentType.Left;

            // Save the workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
