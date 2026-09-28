// Title: Convert an Excel workbook to PDF after programmatically updating SmartArt shape text with Aspose.Cells for .NET (C#)
// AI Prompts: Load an .xlsx file, loop through each Shape on the first worksheet, set its Text property to a new value, and save the workbook as a PDF using Aspose.Cells. | Using Aspose.Cells in C#, replace the text of all SmartArt shapes in a worksheet and export the updated workbook to a PDF file. | Programmatically change shape text in an Excel file and generate a PDF output with Aspose.Cells SaveFormat.Pdf.
// Common Searches: Aspose.Cells C# change SmartArt shape text before PDF conversion | How to iterate over worksheet shapes and modify text with Aspose.Cells .NET | Export Excel to PDF after updating shape text using Aspose.Cells API | C# example for saving modified workbook as PDF with Aspose.Cells | Update SmartArt shapes in Excel and convert to PDF programmatically
// Tags: smartart text replacement aspose.cells | pdf export after shape edit aspose.cells | shape enumeration aspose.cells c# | excel smartart manipulation aspose.cells | c# convert modified workbook to pdf

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The program loads an existing Excel file, iterates through all shapes on the first worksheet to modify their text, and then saves the workbook as a PDF using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Iterate through all shapes on the worksheet
            foreach (Shape shape in sheet.Shapes)
            {
                // Example modification: if the shape contains text, change it
                if (!string.IsNullOrEmpty(shape.Text))
                {
                    shape.Text = "Modified Text";
                }

                // Additional shape modifications can be performed here
            }

            // Export the workbook to PDF
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook successfully saved as PDF: {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
