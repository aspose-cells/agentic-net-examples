// Title: Convert SmartArt to GroupShape and Verify Type with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, adds a placeholder shape, uses GetResultOfSmartArt to turn the SmartArt into a GroupShape, casts the result, checks for null and correct type, and saves the file.
// Keywords: Aspose.Cells | C# | .NET | GetResultOfSmartArt | GroupShape | SmartArt conversion | shape casting | verify shape type | spreadsheet example | Aspose.Cells tutorial
// Common Searches: Aspose.Cells GetResultOfSmartArt GroupShape C# | cast SmartArt result to GroupShape .NET | check SmartArt conversion type Aspose.Cells | validate GroupShape after SmartArt conversion | C# example converting SmartArt to grouped shapes
// Developer Intent: Demonstrate casting the output of GetResultOfSmartArt to GroupShape and confirming the object's type before further manipulation.
// Use Cases: Confirm successful SmartArt‑to‑GroupShape conversion before editing sub‑shapes. | Avoid InvalidCastException by verifying the returned object is a GroupShape. | Apply conditional logic based on shape type when processing imported SmartArt diagrams. | Automate workbook generation that includes SmartArt groups with type safety.
// AI Prompts: Write C# code using Aspose.Cells to convert a SmartArt shape to a GroupShape and assert its type. | Generate a robust example that casts GetResultOfSmartArt output, handles null, and logs verification. | Explain best practices for safely working with a GroupShape returned from SmartArt conversion in Aspose.Cells. | Provide a step‑by‑step guide to validate SmartArt conversion results in a .NET spreadsheet application.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds a placeholder shape, uses GetResultOfSmartArt to turn the SmartArt into a GroupShape, casts the result, checks for null and correct type, and saves the file.
    public class CastSmartArtResultDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape (placeholder for SmartArt)
            Shape shape = worksheet.Shapes.AddRectangle(0, 0, 200, 200, 0, 0);

            // Convert SmartArt to grouped shapes (returns GroupShape)
            GroupShape conversionResult = shape.GetResultOfSmartArt();

            // Verify that the conversion result is a GroupShape
            if (conversionResult != null && conversionResult is GroupShape)
            {
                Console.WriteLine("Conversion returned a GroupShape.");
            }
            else
            {
                Console.WriteLine("Conversion did not return a GroupShape.");
            }

            // Save the workbook
            string outputPath = "CastSmartArtResultDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}
