// Title: Lock a named shape in an Excel workbook to prevent moving or resizing with Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an XLSX file using Aspose.Cells, locates a shape called "Signature" on the first worksheet, sets its IsLocked flag to true, and saves the workbook. | Generate a method that iterates through worksheet shapes in Aspose.Cells for .NET and applies a lock to a specific shape to stop user‑initiated moving or resizing. | Show how to protect a drawing object in an Excel file by configuring the appropriate lock properties with Aspose.Cells in C#.
// Common Searches: Aspose.Cells C# lock shape from being moved in Excel workbook | prevent resizing of a drawing object in XLSX using Aspose.Cells .NET | set IsLocked property on a shape named Signature with Aspose.Cells | how to protect a specific shape in an Excel file using Aspose.Cells for .NET
// Tags: lock shape Aspose.Cells .NET | set IsLocked property Excel drawing object | protect specific shape in XLSX with Aspose.Cells | find shape by name Aspose.Cells C# | prevent shape resizing Aspose.Cells workbook

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook, searches the first worksheet for a shape named "Signature", sets its IsLocked property to true to stop moving or resizing, and saves the modified file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Ensure the input file exists; create a blank workbook if it does not.
            if (!File.Exists(inputPath))
            {
                var newWb = new Workbook();
                newWb.Save(inputPath);
            }

            // Load the workbook.
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed).
            Worksheet worksheet = workbook.Worksheets[0];

            // Locate the shape named "Signature".
            Shape signatureShape = null;
            foreach (Shape shape in worksheet.Shapes)
            {
                if (shape.Name == "Signature")
                {
                    signatureShape = shape;
                    break;
                }
            }

            if (signatureShape != null)
            {
                // Lock the shape to prevent moving and resizing.
                signatureShape.IsLocked = true;            // General lock
                // Additional lock options (if needed) can be set via IsLockedAspectRatio, etc.
            }
            else
            {
                Console.WriteLine("Shape 'Signature' not found.");
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
