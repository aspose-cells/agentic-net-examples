// Title: Resize SmartArt shapes in an Excel worksheet and export the workbook to PDF using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that loads an .xlsx file, locates every SmartArt shape on the first worksheet, updates its Width, Height, Left, and Top properties, and then saves the workbook as a PDF. | Demonstrate how to iterate through Worksheet.Shapes, use Shape.IsSmartArt to identify SmartArt objects, modify their dimensions, and export the modified workbook to PDF with Aspose.Cells for .NET.
// Common Searches: C# Aspose.Cells resize SmartArt before converting Excel to PDF | how to detect and change SmartArt shape size using Aspose.Cells .NET | export modified Excel workbook to PDF after adjusting SmartArt dimensions | Aspose.Cells programmatically reposition SmartArt shapes in a worksheet
// Tags: resize SmartArt shapes Aspose.Cells | export workbook to PDF Aspose.Cells | detect SmartArt programmatically C# | modify shape dimensions Aspose.Cells | adjust SmartArt position .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an existing Excel file, iterates over all shapes on the first worksheet, identifies SmartArt objects via the IsSmartArt property, resizes and repositions each SmartArt, and finally saves the updated workbook as a PDF, with error handling for missing files and runtime exceptions.
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

            // Get the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Iterate through all shapes on the worksheet
            foreach (Shape shape in sheet.Shapes)
            {
                // Identify SmartArt shapes using the IsSmartArt property
                if (shape.IsSmartArt)
                {
                    // Example modification: resize and reposition the SmartArt
                    shape.Width = 300;   // new width in points
                    shape.Height = 200;  // new height in points
                    shape.Left = 50;     // distance from the left edge
                    shape.Top = 50;      // distance from the top edge

                    // Additional modifications can be added here if needed
                }
            }

            // Export the modified workbook to PDF
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook saved as PDF: {outputPath}");
        }
        catch (Exception ex)
        {
            // Log or display the exception details for troubleshooting
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
