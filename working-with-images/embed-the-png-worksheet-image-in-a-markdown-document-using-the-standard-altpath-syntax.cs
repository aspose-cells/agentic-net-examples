// Title: Render the first worksheet of an Excel workbook to a PNG image and embed it in a Markdown file using Aspose.Cells for .NET (C#)
// Common Searches: how to export the first worksheet of an Excel workbook to PNG using Aspose.Cells in C# | C# code to embed a generated worksheet PNG into a markdown document | Aspose.Cells SheetRender ToImage example for creating markdown documentation | set PNG resolution when rendering an Excel sheet with Aspose.Cells | generate markdown with ![alt](path) for an Excel worksheet image in .NET
// Tags: Aspose.Cells SheetRender PNG export | C# export worksheet to PNG image | ImageOrPrintOptions configure PNG resolution | markdown embed Excel worksheet image | Aspose.Cells generate markdown documentation

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing; // Needed for ImageType enum

namespace AsposeCellsMarkdownEmbed
{
    // The sample program loads or creates an Excel workbook, uses Aspose.Cells SheetRender with ImageOrPrintOptions set to PNG (including resolution settings) to render the first worksheet as a PNG file, and then creates a Markdown file that embeds the generated image using the standard ![alt](path) syntax.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Define file names
                const string workbookPath = "SampleWorkbook.xlsx";
                const string imagePath = "WorksheetImage.png";
                const string markdownPath = "WorksheetImage.md";

                // -----------------------------------------------------------------
                // Step 1: Create a sample workbook (or load an existing one)
                // -----------------------------------------------------------------
                Workbook workbook;
                if (File.Exists(workbookPath))
                {
                    // Load existing workbook
                    workbook = new Workbook(workbookPath);
                    Console.WriteLine($"Loaded workbook from \"{workbookPath}\".");
                }
                else
                {
                    // Create a new workbook with some data
                    workbook = new Workbook();
                    Worksheet sheet = workbook.Worksheets[0];
                    sheet.Name = "DemoSheet";

                    // Populate sample data
                    for (int row = 0; row < 10; row++)
                    {
                        for (int col = 0; col < 5; col++)
                        {
                            sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                        }
                    }

                    // Save the workbook for reference (optional)
                    workbook.Save(workbookPath);
                    Console.WriteLine($"Created and saved sample workbook to \"{workbookPath}\".");
                }

                // -----------------------------------------------------------------
                // Step 2: Render the first worksheet to a PNG image
                // -----------------------------------------------------------------
                Worksheet worksheet = workbook.Worksheets[0];

                // Configure image options (PNG format is default)
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    ImageType = ImageType.Png,          // Set output format
                    HorizontalResolution = 150,         // Optional: increase resolution
                    VerticalResolution = 150,
                    Transparent = false
                };

                try
                {
                    // Create SheetRender and render the first page (page index 0)
                    SheetRender renderer = new SheetRender(worksheet, imgOptions);
                    renderer.ToImage(0, imagePath);
                    Console.WriteLine($"Worksheet rendered to PNG image \"{imagePath}\".");
                }
                catch (Exception renderEx)
                {
                    Console.Error.WriteLine($"Error during rendering: {renderEx.Message}");
                    throw; // Re‑throw to be caught by outer handler
                }

                // -----------------------------------------------------------------
                // Step 3: Generate a Markdown file that embeds the PNG image
                // -----------------------------------------------------------------
                string markdownContent = $"# Worksheet Image\n\n" +
                                         $"![Worksheet Image]({imagePath})\n";

                File.WriteAllText(markdownPath, markdownContent);
                Console.WriteLine($"Markdown file created at \"{markdownPath}\" with embedded image.");

                // -----------------------------------------------------------------
                // Completion message
                // -----------------------------------------------------------------
                Console.WriteLine("Process completed successfully.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
                Console.Error.WriteLine(ex.StackTrace);
            }
        }
    }
}
