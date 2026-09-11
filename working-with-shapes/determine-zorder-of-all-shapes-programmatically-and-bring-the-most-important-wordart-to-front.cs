// Title: Programmatically determine shape Z‑order and bring the longest WordArt to front in an Excel worksheet using Aspose.Cells for .NET
// AI Prompts: Generate C# code with Aspose.Cells that iterates over every shape on a worksheet, finds the WordArt shape with the most characters, and sets its ZOrderPosition to move it to the front. | Show how to list all worksheet shapes, filter for WordArt, adjust the layering order via ZOrderPosition, and save the updated workbook using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# get ZOrderPosition of shapes in an Excel worksheet | How to move a specific WordArt shape to the front with Aspose.Cells .NET | Enumerate all shapes and locate the longest WordArt text using Aspose.Cells | Set shape layering order programmatically in an Excel file with Aspose.Cells C# | Bring WordArt to top of stack in Excel via Aspose.Cells API
// Tags: Aspose.Cells shape ZOrderPosition manipulation | C# enumerate worksheet shapes Aspose.Cells | WordArt longest text detection Aspose.Cells | Excel shape layering Aspose.Cells .NET | move WordArt to front Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The example loads an Excel workbook, prints each shape's basic info, identifies the WordArt shape containing the longest text, sets its ZOrderPosition to 0 to bring it to the front, and saves the modified file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one.
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook(); // creates a workbook with a default worksheet
            }

            Worksheet sheet = workbook.Worksheets[0];

            // Output basic info of all shapes on the worksheet.
            foreach (Shape shape in sheet.Shapes)
            {
                Console.WriteLine($"Shape Name: {shape.Name}, Type: {shape.Type}");
            }

            // Find the WordArt shape with the longest text.
            Shape importantWordArt = null;
            int maxTextLength = -1;

            foreach (Shape shape in sheet.Shapes)
            {
                // Detect WordArt by checking the type name (avoids direct enum dependency).
                if (shape.Type.ToString().Contains("WordArt", StringComparison.OrdinalIgnoreCase))
                {
                    string text = shape.Text;
                    if (!string.IsNullOrEmpty(text) && text.Length > maxTextLength)
                    {
                        maxTextLength = text.Length;
                        importantWordArt = shape;
                    }
                }
            }

            // Bring the selected WordArt to the front using ZOrderPosition.
            if (importantWordArt != null)
            {
                // Setting ZOrderPosition to 0 brings the shape to the front.
                importantWordArt.ZOrderPosition = 0;
                Console.WriteLine($"WordArt '{importantWordArt.Name}' brought to front.");
            }
            else
            {
                Console.WriteLine("No WordArt shape found in the worksheet.");
            }

            // Ensure the output directory exists.
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
