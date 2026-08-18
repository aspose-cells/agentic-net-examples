// Title: Extract Excel Workbook Color Palette and Render First Worksheet to TIFF using Aspose.Cells for .NET
// Description: C# example that loads an XLSX file, renders the first worksheet to a TIFF image (LZW compression, 24‑bit), reads the workbook’s full 56‑color palette via Workbook.Colors, retrieves only the colors actually used with CellsHelper.GetUsedColors, prints both collections, and saves the workbook to a new file.
// Keywords: Aspose.Cells | C# | Excel color palette | Workbook.Colors | CellsHelper.GetUsedColors | render worksheet to TIFF | TIFF LZW compression | 56‑color palette | extract used colors | image rendering | .NET Excel processing
// Common Searches: Aspose.Cells get workbook palette C# | list used colors in Excel with Aspose | render Excel sheet to TIFF .NET | extract dominant colors from Excel file | how to read 56‑color palette from workbook
// Developer Intent: Retrieve both the full workbook palette and the subset of colors actually used after converting the first sheet to a TIFF image.
// Use Cases: Create a TIFF snapshot of a worksheet and compare the full 56‑color palette with the colors that are really applied for visual audits. | Identify dominant or theme colors in an Excel file for UI theming or branding analysis. | Validate that custom cell styles use only approved palette entries before publishing a workbook. | Automate generation of color‑usage reports for large Excel datasets.
// AI Prompts: Generate C# code with Aspose.Cells that loads an Excel file, renders the first worksheet to a TIFF, and outputs both the complete 56‑color palette and the used colors. | Explain how CellsHelper.GetUsedColors scans a workbook to determine which palette entries are referenced. | Provide a function that calculates the most frequent color in the used‑colors array and returns its ARGB components.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsColorPaletteExtraction
{
    // C# example that loads an XLSX file, renders the first worksheet to a TIFF image (LZW compression, 24‑bit), reads the workbook’s full 56‑color palette via Workbook.Colors, retrieves only the colors actually used with CellsHelper.GetUsedColors, prints both collections, and saves the workbook to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths for input and output files
            string excelPath = "input.xlsx";
            string tiffPath = "output.tiff";
            string savedWorkbookPath = "modified.xlsx";

            // Verify that the input Excel file exists
            if (!File.Exists(excelPath))
            {
                Console.WriteLine($"Error: Input file \"{excelPath}\" not found.");
                return;
            }

            try
            {
                // Load the workbook (lifecycle rule: load)
                Workbook workbook = new Workbook(excelPath);

                // Ensure there is at least one worksheet to render
                if (workbook.Worksheets.Count == 0)
                {
                    Console.WriteLine("Error: The workbook contains no worksheets.");
                    return;
                }

                // ------------------------------------------------------------
                // Render the first worksheet to a TIFF image file
                // ------------------------------------------------------------
                ImageOrPrintOptions options = new ImageOrPrintOptions
                {
                    // ImageFormat is implicitly TIFF when using ToTiff, so it can be omitted
                    TiffCompression = TiffCompression.CompressionLZW,
                    TiffColorDepth = ColorDepth.Format24bpp,
                    OnePagePerSheet = true
                };

                // Create a SheetRender for the first worksheet
                SheetRender sheetRender = new SheetRender(workbook.Worksheets[0], options);

                // Render all pages of the sheet to the TIFF file (lifecycle rule: save)
                sheetRender.ToTiff(tiffPath);
                Console.WriteLine($"Workbook rendered to TIFF: {tiffPath}");

                // ------------------------------------------------------------
                // Extract the full palette (56 entries) from the workbook
                // ------------------------------------------------------------
                Color[] paletteColors = workbook.Colors; // Returns the 56‑entry palette

                Console.WriteLine("\nFull workbook palette (56 colors):");
                for (int i = 0; i < paletteColors.Length; i++)
                {
                    Color c = paletteColors[i];
                    Console.WriteLine($"Index {i,2}: A={c.A}, R={c.R}, G={c.G}, B={c.B}");
                }

                // ------------------------------------------------------------
                // Extract only the colors that are actually used in the workbook
                // ------------------------------------------------------------
                Color[] usedColors = CellsHelper.GetUsedColors(workbook);

                Console.WriteLine("\nColors actually used in the workbook:");
                foreach (Color c in usedColors)
                {
                    Console.WriteLine($"A={c.A}, R={c.R}, G={c.G}, B={c.B}");
                }

                // ------------------------------------------------------------
                // (Optional) Save the workbook after any modifications
                // ------------------------------------------------------------
                workbook.Save(savedWorkbookPath, SaveFormat.Xlsx);
                Console.WriteLine($"\nWorkbook saved to: {savedWorkbookPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
