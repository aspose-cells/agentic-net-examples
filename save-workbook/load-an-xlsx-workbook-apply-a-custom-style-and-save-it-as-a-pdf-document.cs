// Title: Apply a custom cell style to range A1:C3 in an XLSX workbook and save it as a PDF using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an existing .xlsx file with Aspose.Cells, creates a custom style (Arial 12, blue font, light‑yellow fill, centered alignment), applies it to the range A1:C3, and then exports the workbook to a PDF file. | Show how to use Aspose.Cells Style, StyleFlag, and Range objects to format specific cells and convert the styled worksheet to PDF in a .NET console application.
// Common Searches: aspnet c# apply custom font and background to specific Excel cells with Aspose.Cells then export to PDF | how to format a range in an XLSX file and save as PDF using Aspose.Cells for .NET | Aspose.Cells C# example for styling cells A1:C3 and converting workbook to PDF | create and apply a style to Excel cells before PDF conversion with Aspose.Cells | C# code to load Excel, set cell style, and generate PDF using Aspose.Cells
// Tags: apply custom style Aspose.Cells C# | range formatting Aspose.Cells | export styled worksheet to PDF Aspose.Cells | create Style and StyleFlag Aspose.Cells | convert XLSX to PDF with cell formatting

using Aspose.Cells;
using System;
using System.Drawing;
using System.IO;

// The example checks for the input.xlsx file, loads it with Aspose.Cells, defines a custom style (Arial 12, blue font, light‑yellow background, centered alignment), applies the style to cells A1:C3, and saves the workbook as output.pdf in PDF format, handling any errors that may occur.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the XLSX workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Create a custom style
            Style customStyle = workbook.CreateStyle();
            customStyle.Font.Name = "Arial";
            customStyle.Font.Size = 12;
            customStyle.Font.Color = Color.Blue;
            customStyle.ForegroundColor = Color.LightYellow;
            customStyle.Pattern = BackgroundType.Solid;
            customStyle.HorizontalAlignment = TextAlignmentType.Center;
            customStyle.VerticalAlignment = TextAlignmentType.Center;

            // Apply the style to a specific range (e.g., A1:C3)
            Aspose.Cells.Range targetRange = sheet.Cells.CreateRange("A1:C3");
            StyleFlag flag = new StyleFlag() { All = true };
            targetRange.ApplyStyle(customStyle, flag);

            // Save the workbook as a PDF document
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
