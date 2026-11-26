// Title: Generate a README that documents the PNG worksheet rendering workflow and ImageOrPrintOptions configuration using Aspose.Cells for .NET (C#)
// Common Searches: how to render an Excel worksheet to PNG images with Aspose.Cells in C# and capture the steps in a README | sample code for creating a README that lists PNG conversion options using Aspose.Cells ImageOrPrintOptions | Aspose.Cells ImageOrPrintOptions high‑resolution PNG export example C# | render each page of an Excel sheet as separate PNG files with Aspose.Cells .NET | configure page setup for PNG export (print area, fit to page, orientation) in Aspose.Cells | automate README.md generation after converting a workbook to PNG using Aspose.Cells
// Tags: Aspose.Cells PNG rendering workflow | ImageOrPrintOptions PNG configuration C# | Excel to PNG conversion Aspose.Cells | page setup fit to page Aspose.Cells | generate README after worksheet image export | high DPI PNG output Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing; // For ImageType enum

namespace AsposeCellsPngRender
{
    // The C# program loads or creates an Excel workbook, sets page‑setup properties (print area, fit‑to‑page, portrait orientation), configures ImageOrPrintOptions for PNG output (image type, one page per sheet, 300 dpi), renders each worksheet page to separate PNG files in a 'PngOutput' folder, and automatically creates a README.md that records the input file, page‑setup values, PNG options, page count, and generated file names.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths (relative to the executable folder)
            string inputPath = "input.xlsx";
            string outputFolder = "PngOutput";
            string readmePath = Path.Combine(outputFolder, "README.md");

            try
            {
                // Ensure output directory exists
                Directory.CreateDirectory(outputFolder);

                // ------------------------------------------------------------
                // 1. Load or create a workbook
                // ------------------------------------------------------------
                Workbook workbook;
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                    Console.WriteLine($"Loaded workbook from \"{inputPath}\".");
                }
                else
                {
                    // Create a simple workbook with sample data if the file does not exist
                    workbook = new Workbook();
                    Worksheet sheet = workbook.Worksheets[0];
                    sheet.Name = "SampleData";

                    // Populate sample data
                    for (int row = 0; row < 20; row++)
                    {
                        for (int col = 0; col < 5; col++)
                        {
                            sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                        }
                    }

                    // Save the generated workbook for reference
                    workbook.Save(inputPath);
                    Console.WriteLine($"Created sample workbook and saved to \"{inputPath}\".");
                }

                // ------------------------------------------------------------
                // 2. Configure page setup (optional but often required for PNG rendering)
                // ------------------------------------------------------------
                Worksheet worksheet = workbook.Worksheets[0];
                PageSetup pageSetup = worksheet.PageSetup;
                // Example: print area covers rows 1-20 and columns A-E
                pageSetup.PrintArea = "A1:E20";
                // Fit the whole sheet onto a single page
                pageSetup.SetFitToPages(1, 1);
                // Set orientation (use PageOrientationType enum)
                pageSetup.Orientation = PageOrientationType.Portrait;

                // ------------------------------------------------------------
                // 3. Define ImageOrPrintOptions for PNG output
                // ------------------------------------------------------------
                ImageOrPrintOptions options = new ImageOrPrintOptions
                {
                    // Set the output image type to PNG
                    ImageType = ImageType.Png,
                    // Render each worksheet page as a separate image file
                    OnePagePerSheet = true,
                    // High resolution for better quality (dots per inch)
                    HorizontalResolution = 300,
                    VerticalResolution = 300,
                    // Optional: set background color (transparent by default for PNG)
                    // BackgroundColor = System.Drawing.Color.Transparent,
                };

                // ------------------------------------------------------------
                // 4. Create SheetRender and render each page to PNG
                // ------------------------------------------------------------
                SheetRender renderer = new SheetRender(worksheet, options);
                int pageCount = renderer.PageCount;
                Console.WriteLine($"Worksheet \"{worksheet.Name}\" will be rendered to {pageCount} PNG page(s).");

                // Store generated file names for README
                string[] generatedFiles = new string[pageCount];

                for (int pageIndex = 0; pageIndex < pageCount; pageIndex++)
                {
                    string pngFileName = Path.Combine(outputFolder, $"Page_{pageIndex + 1}.png");
                    renderer.ToImage(pageIndex, pngFileName);
                    generatedFiles[pageIndex] = Path.GetFileName(pngFileName);
                    Console.WriteLine($"Rendered page {pageIndex + 1} to \"{pngFileName}\".");
                }

                // ------------------------------------------------------------
                // 5. Generate README.md documenting the conversion steps and options
                // ------------------------------------------------------------
                using (StreamWriter writer = new StreamWriter(readmePath, false))
                {
                    writer.WriteLine("# PNG Worksheet Rendering with Aspose.Cells for .NET");
                    writer.WriteLine();
                    writer.WriteLine("## Overview");
                    writer.WriteLine();
                    writer.WriteLine("This README documents the exact steps and **ImageOrPrintOptions** used to convert a worksheet to PNG images using the Aspose.Cells for .NET library.");
                    writer.WriteLine();
                    writer.WriteLine("## Input");
                    writer.WriteLine();
                    writer.WriteLine($"- Workbook file: `{Path.GetFullPath(inputPath)}`");
                    writer.WriteLine("- Worksheet rendered: first worksheet (`{worksheet.Name}`)");
                    writer.WriteLine();
                    writer.WriteLine("## Page Setup Configuration");
                    writer.WriteLine();
                    writer.WriteLine("| Property | Value |");
                    writer.WriteLine("|----------|-------|");
                    writer.WriteLine($"| PrintArea | `{pageSetup.PrintArea}` |");
                    writer.WriteLine($"| FitToPagesWide | `{pageSetup.FitToPagesWide}` |");
                    writer.WriteLine($"| FitToPagesTall | `{pageSetup.FitToPagesTall}` |");
                    writer.WriteLine($"| Orientation | `{pageSetup.Orientation}` |");
                    writer.WriteLine();
                    writer.WriteLine("## ImageOrPrintOptions Used");
                    writer.WriteLine();
                    writer.WriteLine("| Option | Value |");
                    writer.WriteLine("|--------|-------|");
                    writer.WriteLine($"| ImageType | `Png` |");
                    writer.WriteLine($"| OnePagePerSheet | `true` |");
                    writer.WriteLine($"| HorizontalResolution | `{options.HorizontalResolution}` DPI |");
                    writer.WriteLine($"| VerticalResolution | `{options.VerticalResolution}` DPI |");
                    writer.WriteLine();
                    writer.WriteLine("## Rendering Result");
                    writer.WriteLine();
                    writer.WriteLine($"- Total pages rendered: **{pageCount}**");
                    writer.WriteLine("- Generated PNG files:");
                    writer.WriteLine();
                    foreach (string file in generatedFiles)
                    {
                        writer.WriteLine($"  - `{file}`");
                    }
                    writer.WriteLine();
                    writer.WriteLine("All PNG files are stored in the `PngOutput` directory next to the executable.");
                    writer.WriteLine();
                    writer.WriteLine("## Code Snippet");
                    writer.WriteLine();
                    writer.WriteLine("```csharp");
                    writer.WriteLine("using Aspose.Cells;");
                    writer.WriteLine("using Aspose.Cells.Rendering;");
                    writer.WriteLine("// ... (code shown in this Program.cs) ...");
                    writer.WriteLine("```");
                }

                Console.WriteLine($"README documentation generated at \"{readmePath}\".");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("An error occurred during PNG rendering:");
                Console.Error.WriteLine(ex.ToString());
                Environment.Exit(1);
            }
        }
    }
}
