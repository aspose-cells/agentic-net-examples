// Title: How to replace an Excel header WordArt with a preset style using Aspose.Cells for .NET (C#)
// Common Searches: C# code to change WordArt style in an existing .xlsx file with Aspose.Cells | Aspose.Cells replace WordArt shape on Excel worksheet programmatically | Update header WordArt font size and make it bold in Excel using Aspose.Cells .NET | Iterate through shapes in Aspose.Cells to modify WordArt properties | Apply PresetWordArtStyle.WordArtStyle7 to an Excel header using Aspose.Cells C#
// Tags: modify WordArt style Aspose.Cells | replace header WordArt Excel C# | apply PresetWordArtStyle WordArtStyle7 | iterate worksheet shapes Aspose.Cells | set WordArt font size bold

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

namespace AsposeCellsWordArtReplace
{
    // The example loads an existing workbook, iterates over each worksheet and its shapes, identifies WordArt shapes, applies the PresetWordArtStyle.WordArtStyle7 preset, sets the font size to 24 points and makes the text bold, then saves the updated workbook to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths to the input and output Excel files.
            // Ensure that the input file exists at the specified location.
            string inputFile = "HeaderInput.xlsx";
            string outputFile = "HeaderOutput.xlsx";

            try
            {
                // Load the existing workbook.
                Workbook workbook = new Workbook(inputFile);

                // Iterate through all worksheets (or limit to a specific one if required).
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through all shapes on the worksheet.
                    foreach (Shape shape in sheet.Shapes)
                    {
                        // Identify WordArt shapes using the IsWordArt property.
                        if (shape.IsWordArt)
                        {
                            Console.WriteLine($"Found WordArt shape on sheet \"{sheet.Name}\" with text: \"{shape.Text}\"");

                            // Apply a new preset WordArt style.
                            // Example: using WordArtStyle7 (Gradient Fill - Blue, Accent 1, Reflection).
                            shape.TextBody.SetWordArtStyle(PresetWordArtStyle.WordArtStyle7);
                            Console.WriteLine("Applied PresetWordArtStyle.WordArtStyle7.");

                            // Optionally adjust additional text effect properties such as font size.
                            // TextEffect is available for WordArt shapes.
                            shape.TextEffect.FontSize = 24; // Set font size to 24 points.
                            shape.TextEffect.FontBold = true;
                            Console.WriteLine("Set FontSize to 24 points and made it bold.");

                            // If you need to change the actual text, you can do so:
                            // shape.Text = "New Header Text";
                        }
                    }
                }

                // Save the modified workbook.
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved successfully to \"{outputFile}\".");
            }
            catch (Exception ex)
            {
                // Report any errors that occur during processing.
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
