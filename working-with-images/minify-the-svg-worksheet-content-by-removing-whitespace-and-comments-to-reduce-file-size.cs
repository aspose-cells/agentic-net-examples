// Title: C# program to minify Aspose.Cells‑generated worksheet SVG by removing comments and excess whitespace
// Common Searches: C# remove XML comments from SVG produced by Aspose.Cells | How to reduce size of worksheet SVG using Aspose.Cells in .NET | Minify SVG output from Aspose.Cells rendering with regular expressions | Strip whitespace from Aspose.Cells SVG file in a console application | Save a compressed SVG version of an Excel sheet using Aspose.Cells
// Tags: Aspose.Cells SVG rendering options | C# regex based SVG minifier | worksheet SVG compression .NET | memory stream handling for SVG output | strip XML comments in SVG data

using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace SvgMinifier
{
    // The example loads an Excel workbook, renders the first worksheet to SVG with Aspose.Cells, reads the SVG from a memory stream, then uses regular expressions to strip XML comments and collapse whitespace, producing a smaller SVG file saved to the specified path.
    class Program
    {
        static void Main(string[] args)
        {
            // Validate arguments
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: SvgMinifier <input.xlsx> <output.svg>");
                return;
            }

            string inputPath = args[0];
            string outputPath = args[1];

            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" does not exist.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);
                Worksheet sheet = workbook.Worksheets[0]; // Process the first worksheet

                // Configure rendering options for SVG output
                ImageOrPrintOptions options = new ImageOrPrintOptions
                {
                    ImageType = Aspose.Cells.Drawing.ImageType.Svg,
                    OnePagePerSheet = true,
                    // Optional: set a CSS prefix to avoid style collisions (not required for minification)
                    // SvgCssPrefix = "c_"
                };

                // Render the worksheet to SVG using a memory stream
                using (MemoryStream svgStream = new MemoryStream())
                {
                    SheetRender renderer = new SheetRender(sheet, options);
                    // Render the first (and only) page
                    renderer.ToImage(0, svgStream);
                    svgStream.Position = 0;

                    // Read the raw SVG content
                    string rawSvg;
                    using (StreamReader reader = new StreamReader(svgStream, Encoding.UTF8))
                    {
                        rawSvg = reader.ReadToEnd();
                    }

                    // Minify the SVG: remove XML comments and unnecessary whitespace
                    string minifiedSvg = MinifySvg(rawSvg);

                    // Write the minified SVG to the output file
                    File.WriteAllText(outputPath, minifiedSvg, Encoding.UTF8);
                    Console.WriteLine($"Minified SVG saved to \"{outputPath}\".");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred during processing:");
                Console.WriteLine(ex.Message);
            }
        }

        /// <summary>
        /// Removes XML comments and collapses whitespace in an SVG string.
        /// </summary>
        /// <param name="svg">The original SVG content.</param>
        /// <returns>Minified SVG content.</returns>
        private static string MinifySvg(string svg)
        {
            if (string.IsNullOrEmpty(svg))
                return svg;

            // Remove XML comments: <!-- comment -->
            string withoutComments = Regex.Replace(svg, @"<!--(.*?)-->", string.Empty, RegexOptions.Singleline);

            // Collapse multiple whitespace characters into a single space
            string collapsedWhitespace = Regex.Replace(withoutComments, @"\s+", " ");

            // Remove spaces between tags: > <  => ><
            string trimmedBetweenTags = Regex.Replace(collapsedWhitespace, @">\s+<", "><");

            // Trim leading/trailing whitespace
            return trimmedBetweenTags.Trim();
        }
    }
}
