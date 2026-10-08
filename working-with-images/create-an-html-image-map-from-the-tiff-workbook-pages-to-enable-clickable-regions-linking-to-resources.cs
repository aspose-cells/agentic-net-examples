// Title: Generate an HTML image map with clickable cells from Excel worksheets rendered as PNG using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an Excel workbook with Aspose.Cells, renders each worksheet to a PNG image, and builds an HTML <map> containing <area> elements for every cell that has a hyperlink. | Implement a method that calculates the pixel rectangle of a specific cell based on column widths and row heights, then uses those coordinates to define rectangular clickable regions in an HTML image map. | Create a single HTML file that embeds the rendered worksheet images with usemap attributes, linking each image to its corresponding map of hyperlink areas.
// Common Searches: how to generate an HTML image map from Excel sheets using Aspose.Cells C# | Aspose.Cells render worksheet to PNG and create clickable regions for cell hyperlinks | C# calculate Excel cell pixel coordinates for image map generation | export Excel workbook as images with hyperlink areas in HTML | build HTML <map> from Excel hyperlinks using Aspose.Cells .NET
// Tags: Aspose.Cells render worksheet to PNG | HTML image map generation from Excel cells | calculate cell pixel rectangle C# | hyperlink area tags for Excel image map | export Excel as images with clickable regions | C# image map creation using Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example loads an Excel workbook, renders each worksheet to a PNG image with Aspose.Cells, computes pixel rectangles for cells containing hyperlinks, constructs an HTML <map> with <area> elements using those coordinates, and writes a single HTML file that displays the images linked to their respective clickable regions.
class TiffImageMapGenerator
{
    static void Main()
    {
        try
        {
            // Input workbook path
            string inputPath = "input.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Prepare output directories
            string outputDir = "output";
            string imageDir = Path.Combine(outputDir, "Images");
            Directory.CreateDirectory(imageDir);
            Directory.CreateDirectory(outputDir);

            // HTML builder
            StringBuilder htmlBuilder = new StringBuilder();
            htmlBuilder.AppendLine("<!DOCTYPE html>");
            htmlBuilder.AppendLine("<html><head><meta charset=\"UTF-8\"><title>Image Map</title></head><body>");

            // Options for image rendering (default PNG)
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                HorizontalResolution = 96,
                VerticalResolution = 96,
                OnePagePerSheet = true
            };

            // Iterate through each worksheet
            for (int sheetIndex = 0; sheetIndex < workbook.Worksheets.Count; sheetIndex++)
            {
                try
                {
                    Worksheet sheet = workbook.Worksheets[sheetIndex];
                    string sheetName = sheet.Name;

                    // ---------- 1. Render the worksheet as an image ----------
                    string imgFileName = $"{sheetName}.png";
                    string imgPath = Path.Combine(imageDir, imgFileName);

                    // Ensure full sheet is printed
                    sheet.PageSetup.PrintArea = "";
                    sheet.PageSetup.PrintTitleRows = "";
                    sheet.PageSetup.PrintTitleColumns = "";

                    // Render the sheet to a single-page image
                    SheetRender renderer = new SheetRender(sheet, imgOptions);
                    renderer.ToImage(0, imgPath);

                    // ---------- 2. Build image map based on cells that contain hyperlinks ----------
                    StringBuilder mapBuilder = new StringBuilder();
                    string mapName = $"map_{sheetIndex}";
                    mapBuilder.AppendLine($"<map name=\"{mapName}\">");

                    foreach (Hyperlink hl in sheet.Hyperlinks)
                    {
                        // Only process external hyperlinks (address not empty)
                        if (string.IsNullOrEmpty(hl.Address))
                            continue;

                        int row = hl.Area.StartRow;
                        int col = hl.Area.StartColumn;

                        // Compute pixel rectangle for the cell
                        Rectangle rect = GetCellRectangle(sheet, row, col);

                        // Build the <area> element
                        string area = $"<area shape=\"rect\" coords=\"{rect.Left},{rect.Top},{rect.Right},{rect.Bottom}\" href=\"{hl.Address}\" alt=\"{sheet.Cells[row, col].StringValue}\" />";
                        mapBuilder.AppendLine(area);
                    }

                    mapBuilder.AppendLine("</map>");

                    // ---------- 3. Insert the image and its map into the HTML ----------
                    string relativeImagePath = Path.Combine("Images", imgFileName).Replace("\\", "/");
                    htmlBuilder.AppendLine($"<h2>{sheetName}</h2>");
                    htmlBuilder.AppendLine($"<img src=\"{relativeImagePath}\" usemap=\"#{mapName}\" alt=\"{sheetName}\" />");
                    htmlBuilder.AppendLine(mapBuilder.ToString());
                }
                catch (Exception sheetEx)
                {
                    Console.WriteLine($"Error processing sheet index {sheetIndex}: {sheetEx.Message}");
                }
            }

            htmlBuilder.AppendLine("</body></html>");

            // Save the final HTML file
            string htmlPath = Path.Combine(outputDir, "WorkbookImageMap.html");
            File.WriteAllText(htmlPath, htmlBuilder.ToString());

            Console.WriteLine("Images and HTML image map have been generated successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Calculates the pixel rectangle of a cell based on column widths and row heights.
    private static Rectangle GetCellRectangle(Worksheet sheet, int row, int col)
    {
        // Left offset: sum of widths of all preceding columns
        int left = 0;
        for (int c = 0; c < col; c++)
        {
            left += sheet.Cells.GetColumnWidthPixel(c);
        }

        // Top offset: sum of heights of all preceding rows (convert points to pixels)
        int top = 0;
        for (int r = 0; r < row; r++)
        {
            double heightPoints = sheet.Cells.GetRowHeight(r);
            int heightPixels = (int)Math.Round(heightPoints * 96.0 / 72.0);
            top += heightPixels;
        }

        // Width and height of the current cell
        int width = sheet.Cells.GetColumnWidthPixel(col);
        double heightPts = sheet.Cells.GetRowHeight(row);
        int height = (int)Math.Round(heightPts * 96.0 / 72.0);

        return new Rectangle(left, top, width, height);
    }
}
