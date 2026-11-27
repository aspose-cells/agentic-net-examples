// Title: How to change the text color of all WordArt shapes to dark blue in an Excel workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# change WordArt font color to dark blue in all worksheets | iterate through Excel shapes and set WordArt text color using Aspose.Cells .NET | C# code to update WordArt objects color in a workbook with Aspose.Cells | bulk modify WordArt text color in Excel file using Aspose.Cells library | set Font.Color for WordArt shapes across multiple sheets Aspose.Cells C#
// Tags: Aspose.Cells modify WordArt font color | C# iterate worksheet shape collection | bulk update WordArt text style in Excel | set Font.Color for WordArt using Aspose.Cells | process Excel shapes with Aspose.Cells .NET

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace WordArtColorChanger
{
    // The example loads an Excel workbook, iterates each worksheet's ShapeCollection, detects WordArt objects via IsWordArt, changes their Font.Color to Color.DarkBlue, and saves the updated file.
    class Program
    {
        /// <param name="args">Command‑line arguments.</param>
        static void Main(string[] args)
        {
            // Validate arguments
            if (args == null || args.Length < 2)
            {
                Console.WriteLine("Usage: WordArtColorChanger <input.xlsx> <output.xlsx>");
                return;
            }

            string inputPath = args[0];
            string outputPath = args[1];

            // Verify that the input file exists
            if (!System.IO.File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" does not exist.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through each worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Get the collection of shapes on the current worksheet
                    ShapeCollection shapes = sheet.Shapes;

                    // Iterate over all shapes
                    foreach (Shape shape in shapes)
                    {
                        // Identify WordArt objects using the IsWordArt property
                        if (shape.IsWordArt)
                        {
                            // Change the text color to dark blue.
                            // For WordArt, the Font property controls the text appearance.
                            shape.Font.Color = Color.DarkBlue;

                            // Optionally, you can also adjust the TextEffect format if needed.
                            // shape.TextEffect.FontColor = Color.DarkBlue; // Uncomment if required.
                        }
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Successfully processed \"{inputPath}\" and saved result to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Output any unexpected errors
                Console.WriteLine("An error occurred while processing the workbook:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
