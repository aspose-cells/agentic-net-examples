// Title: Read and Change Shape Z‑Order, Bring WordArt to Front with Aspose.Cells for .NET (C#)
// Description: This example loads an Excel file, enumerates every shape on the first worksheet, prints each shape's ZOrderPosition, identifies the first WordArt shape by name, moves it to the top layer using the ZOrderPosition property or ToFrontOrBack method, and saves the updated workbook.
// Keywords: Aspose.Cells C# shape Z-order | move WordArt to front Aspose.Cells | Worksheet.Shapes enumeration | ZOrderPosition property | ToFrontOrBack method | Excel shape layering .NET | Aspose.Cells Drawing API
// Common Searches: how to get shape Z-order in Aspose.Cells | bring WordArt to front programmatically C# | set ZOrderPosition of a shape Aspose.Cells | list all shapes on a worksheet Aspose.Cells | change shape layering in Excel using .NET
// Developer Intent: Read the Z‑order of every shape in a worksheet and promote the most important WordArt shape to the front layer.
// Use Cases: Debug visual stacking by printing each shape's ZOrderPosition. | Select a WordArt shape based on its name and elevate it for emphasis. | Persist the new visual hierarchy by saving the workbook after reordering shapes.
// AI Prompts: Write C# code with Aspose.Cells that lists all worksheet shapes and their ZOrderPosition, then moves a shape whose name contains "WordArt" to the front. | Show how to use the ToFrontOrBack method instead of setting ZOrderPosition directly to bring a shape forward. | Explain how to sort shapes by ZOrderPosition and assign a specific shape as the topmost layer in an Excel file.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsZOrderDemo
{
    // This example loads an Excel file, enumerates every shape on the first worksheet, prints each shape's ZOrderPosition, identifies the first WordArt shape by name, moves it to the top layer using the ZOrderPosition property or ToFrontOrBack method, and saves the updated workbook.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your file path)
            Workbook workbook = new Workbook("Input.xlsx");
            Worksheet sheet = workbook.Worksheets[0];

            // Iterate through all shapes to display their Z‑order positions
            Console.WriteLine("Current Z‑order of shapes:");
            foreach (Shape shape in sheet.Shapes)
            {
                Console.WriteLine($"Shape Name: {shape.Name}, ZOrderPosition: {shape.ZOrderPosition}");
            }

            // Find the most important WordArt shape.
            // Here we assume WordArt shapes have "WordArt" in their name.
            Shape mostImportantWordArt = null;
            foreach (Shape shape in sheet.Shapes)
            {
                if (shape.Name != null && shape.Name.IndexOf("WordArt", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    // Choose the first WordArt found as the most important.
                    // You can replace this logic with your own importance criteria.
                    mostImportantWordArt = shape;
                    break;
                }
            }

            if (mostImportantWordArt != null)
            {
                // Bring the selected WordArt to the front.
                // Setting ZOrderPosition to 0 makes it the frontmost shape.
                mostImportantWordArt.ZOrderPosition = 0;

                // Alternatively, you can use ToFrontOrBack with a positive value.
                // mostImportantWordArt.ToFrontOrBack(1);
                Console.WriteLine($"WordArt '{mostImportantWordArt.Name}' moved to front.");
            }
            else
            {
                Console.WriteLine("No WordArt shape found in the worksheet.");
            }

            // Save the modified workbook (replace with your desired output path)
            workbook.Save("Output.xlsx");
        }
    }
}
