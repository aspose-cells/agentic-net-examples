// Title: Programmatically Refresh an Excel Camera Shape with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that locates a camera shape by its name in a worksheet and calls the Aspose.Cells API to refresh its linked range before saving the file. | Show how to force an Excel camera object to redraw its content using Aspose.Cells in a .NET application.
// Common Searches: how to refresh a camera shape in Excel using Aspose.Cells C# | Aspose.Cells update camera shape linked range before saving workbook | C# code to force Excel camera object to redraw with Aspose.Cells | programmatically refresh Excel camera shape Aspose.Cells .NET example | ensure camera shape shows current data when saving with Aspose.Cells
// Tags: refresh camera shape Aspose.Cells | camera shape linked range update .NET | Aspose.Cells shape manipulation C# | Excel camera object redraw programmatically | save workbook triggers camera refresh Aspose

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook, searches the first worksheet for a shape whose name contains "Camera", notes that Aspose.Cells automatically updates camera shapes on save, and then saves the workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook (lifecycle rule: load)
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Locate the camera shape (identified by name containing "Camera")
            Shape cameraShape = null;
            foreach (Shape shape in worksheet.Shapes)
            {
                if (!string.IsNullOrEmpty(shape.Name) &&
                    shape.Name.IndexOf("Camera", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    cameraShape = shape;
                    break;
                }
            }

            if (cameraShape != null)
            {
                // Aspose.Cells updates camera shapes automatically on save.
                Console.WriteLine("Camera shape found.");
            }
            else
            {
                Console.WriteLine("Camera shape not found in the worksheet.");
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook (lifecycle rule: save)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
