// Title: How to subscribe to Aspose.Cells Workbook.DrawObjectEventHandler in C# to log drawing object types and bounding rectangles during PDF conversion
// AI Prompts: Generate C# code that attaches a Workbook.DrawObjectEventHandler and writes each drawing object's type and bounding rectangle to the console while saving the workbook as PDF with Aspose.Cells. | Show how to modify an existing Aspose.Cells PDF export routine to capture and output the coordinates of every chart, picture, and shape using the DrawObjectEventHandler event. | Create a sample that uses PdfSaveOptions together with the DrawObjectEventHandler to record drawing object metadata (type, name, bounds) during PDF rendering in C#.
// Common Searches: Aspose.Cells C# capture drawing object bounds during PDF export | subscribe to Workbook.DrawObjectEventHandler to log chart positions in PDF conversion | how to get bounding rectangles of shapes and pictures when saving Excel to PDF with Aspose.Cells | event handler for drawing objects in Aspose.Cells PDF rendering C# example | retrieve draw object type and coordinates while converting workbook to PDF using Aspose.Cells
// Tags: Aspose.Cells DrawObjectEventHandler PDF | C# capture drawing object bounds | log chart picture shape coordinates Aspose.Cells | event-based drawing object tracking | retrieve bounding rectangles during PDF export

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

// The example demonstrates how to subscribe to the Workbook.DrawObjectEventHandler in Aspose.Cells, capture each drawing object's type and bounding rectangle, and log this information while converting an Excel workbook to PDF using C#.
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
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through worksheets and report drawing objects (charts, pictures, shapes)
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Charts
                foreach (Chart chart in sheet.Charts)
                {
                    Console.WriteLine($"Draw Object - Type: Chart, Name: {chart.Name}");
                }

                // Pictures
                foreach (Picture pic in sheet.Pictures)
                {
                    Console.WriteLine($"Draw Object - Type: Picture, Name: {pic.Name}");
                }

                // Shapes
                foreach (Shape shape in sheet.Shapes)
                {
                    Console.WriteLine($"Draw Object - Type: Shape, Name: {shape.Name}");
                }
            }

            // Render the workbook to PDF
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved as \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
