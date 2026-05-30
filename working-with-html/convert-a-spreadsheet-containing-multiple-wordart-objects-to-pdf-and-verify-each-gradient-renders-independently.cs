using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsWordArtGradientPdf
{
    class Program
    {
        static void Main()
        {
            // Load the workbook that already contains multiple WordArt objects.
            // Replace "input.xlsx" with the actual path to your source spreadsheet.
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet (adjust if needed).
            Worksheet worksheet = workbook.Worksheets[0];

            // Iterate through all shapes in the worksheet.
            foreach (Shape shape in worksheet.Shapes)
            {
                // Process only WordArt shapes.
                if (shape.IsWordArt)
                {
                    Console.WriteLine($"Found WordArt: \"{shape.TextEffect.Text}\"");

                    // Check if the WordArt uses a gradient fill.
                    if (shape.Fill.FillType == FillType.Gradient)
                    {
                        // Obtain the GradientFill object.
                        GradientFill gradientFill = shape.Fill.GradientFill;

                        if (gradientFill != null)
                        {
                            // Output basic gradient information.
                            Console.WriteLine("  Gradient Fill detected.");
                            Console.WriteLine($"  Gradient Style: {gradientFill.GetType().Name}");
                            // GradientStops collection holds the individual colors; display count.
                            Console.WriteLine($"  Number of Gradient Stops: {gradientFill.GradientStops.Count}");
                        }
                        else
                        {
                            Console.WriteLine("  Gradient Fill is null – unexpected state.");
                        }
                    }
                    else
                    {
                        Console.WriteLine("  No gradient fill applied to this WordArt.");
                    }
                }
            }

            // Save the workbook as PDF. Each WordArt gradient should render independently in the PDF.
            // Replace "output.pdf" with the desired output path.
            workbook.Save("output.pdf", SaveFormat.Pdf);

            Console.WriteLine("Conversion to PDF completed.");
        }
    }
}