// Title: C# – Render Multiple Worksheets to PNG Frames for Animated GIF with Aspose.Cells
// Description: This example creates a workbook, adds five sequential worksheets, fills each with time‑step data, renders every sheet to a PNG image using Aspose.Cells SheetRender, and stores the frames in a "Frames" folder. The sample stops before GIF assembly, noting the need for additional graphics libraries.
// Keywords: Aspose.Cells C# render worksheet PNG | export Excel sheets as images .NET | create PNG frames from workbook | animated GIF from Excel images | time‑step worksheet snapshots | SheetRender example | C# image sequence generation
// Common Searches: how to export each Excel worksheet to PNG using Aspose.Cells | C# generate PNG frames from multiple sheets for GIF | Aspose.Cells render workbook to image sequence | save Excel worksheets as PNG files programmatically | build animated GIF from Excel sheet images C#
// Developer Intent: Generate PNG images for every worksheet so they can be combined into an animated GIF.
// Use Cases: Create frame‑by‑frame visualizations of simulation results stored in separate worksheets. | Produce a GIF that shows monthly financial KPI changes by rendering each month’s sheet to PNG. | Automate step‑by‑step tutorial screenshots from a workbook for documentation or training videos.
// AI Prompts: Write C# code that reads the PNG files in the "Frames" folder and assembles them into an animated GIF using System.Drawing or a modern library like ImageSharp. | Show how to adjust the PNG rendering options to set DPI, background color, and image size for each frame. | Explain how to add fallback handling when the required graphics library for GIF creation is unavailable.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsAnimatedGifDemo
{
    // This example creates a workbook, adds five sequential worksheets, fills each with time‑step data, renders every sheet to a PNG image using Aspose.Cells SheetRender, and stores the frames in a "Frames" folder. The sample stops before GIF assembly, noting the need for additional graphics libraries.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Generate several worksheets representing different time steps
                int timeSteps = 5;
                for (int i = 0; i < timeSteps; i++)
                {
                    Worksheet sheet;
                    if (i == 0)
                    {
                        // First worksheet already exists
                        sheet = workbook.Worksheets[0];
                    }
                    else
                    {
                        // Add a new worksheet and obtain its reference
                        int newIndex = workbook.Worksheets.Add();
                        sheet = workbook.Worksheets[newIndex];
                    }

                    sheet.Name = $"Step_{i + 1}";

                    // Fill the sheet with sample data that changes over time
                    sheet.Cells["A1"].PutValue("Time Step");
                    sheet.Cells["B1"].PutValue(i + 1);
                    sheet.Cells["A2"].PutValue("Value");
                    sheet.Cells["B2"].PutValue((i + 1) * 10);
                }

                // Ensure the output folder exists
                string outputFolder = "Frames";
                if (!Directory.Exists(outputFolder))
                {
                    Directory.CreateDirectory(outputFolder);
                }

                // Render each worksheet to a PNG image
                int sheetIndex = 0;
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    try
                    {
                        ImageOrPrintOptions options = new ImageOrPrintOptions
                        {
                            ImageType = Aspose.Cells.Drawing.ImageType.Png,
                            OnePagePerSheet = true
                        };

                        SheetRender renderer = new SheetRender(sheet, options);
                        using (MemoryStream ms = new MemoryStream())
                        {
                            renderer.ToImage(0, ms);
                            string filePath = Path.Combine(outputFolder, $"frame_{sheetIndex + 1}.png");
                            File.WriteAllBytes(filePath, ms.ToArray());
                        }
                        renderer.Dispose();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to render sheet '{sheet.Name}': {ex.Message}");
                    }

                    sheetIndex++;
                }

                Console.WriteLine($"PNG frames saved to folder: {Path.GetFullPath(outputFolder)}");
                Console.WriteLine("Animated GIF creation is omitted due to missing graphics dependencies.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
