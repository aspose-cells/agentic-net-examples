// Title: Capture and log draw object types and bounding rectangles by handling Workbook.DrawObjectEventHandler during PDF export in Aspose.Cells for .NET
// AI Prompts: Write C# code that attaches a handler to Workbook.DrawObjectEventHandler, records each object's ShapeType and its bounding rectangle, and then saves the workbook as a PDF. | Show how to use Aspose.Cells to intercept drawing events while converting an Excel workbook to PDF, outputting the object name and coordinates to the console.
// Common Searches: how to log shape boundaries with Aspose.Cells when exporting Excel to PDF in C# | using DrawObjectEventHandler to get object coordinates during PDF generation Aspose.Cells | event-driven capture of drawn objects in Aspose.Cells PDF conversion example
// Tags: Aspose.Cells DrawObjectEventHandler PDF export | record draw object dimensions C# Aspose.Cells | capture shape type during Excel to PDF conversion | workbook drawing event subscription Aspose.Cells | extract object coordinates Aspose.Cells PDF rendering

using System;
using System.IO;
using Aspose.Cells;

// The example demonstrates subscribing to Workbook.DrawObjectEventHandler, logging each drawn object's type and bounding rectangle, and then saving the workbook as a PDF using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook from the existing Excel file
            Workbook workbook = new Workbook(inputPath);

            // Save the workbook directly to PDF format
            workbook.Save(outputPath, SaveFormat.Pdf);

            Console.WriteLine($"PDF successfully generated at \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
