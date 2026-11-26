// Title: Generate an HTML image map from Excel worksheets rendered as TIFF images with clickable comment links using Aspose.Cells for .NET (C#)
// Common Searches: how to create an HTML usemap from Excel worksheet images with Aspose.Cells C# | Aspose.Cells render worksheet to TIFF and map comment URLs to clickable areas | calculate cell pixel coordinates for HTML image map from Excel comments in C# | export Excel comment notes as hyperlink regions in an image map using Aspose.Cells | C# program to generate HTML image map from multiple worksheets TIFF files
// Tags: render worksheet to TIFF Aspose.Cells | generate HTML usemap from Excel comments | calculate cell pixel coordinates for image map | export comment notes as hyperlink areas | SheetRender ToImage single page TIFF

using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The C# utility validates input paths, loads an Excel workbook, and iterates through each worksheet. For every sheet it renders a single‑page TIFF image using Aspose.Cells' SheetRender, saves the image, and builds an HTML file containing an <img> tag with a usemap attribute. It scans the worksheet's comments, treats each comment's Note as a URL, computes the pixel rectangle of the anchored cell using column widths and row heights, and creates corresponding <area> elements inside a <map>. If no comment URLs exist, a placeholder comment is added. The final HTML image map is written to the specified output folder.
class Program
{
    static void Main(string[] args)
    {
        // ------------------------------------------------------------
        // 1. Validate arguments
        // ------------------------------------------------------------
        if (args.Length < 2)
        {
            Console.WriteLine("Usage: Program.exe <inputWorkbookPath> <outputFolder>");
            return;
        }

        string workbookPath = args[0];
        string outputFolder = args[1];

        if (!File.Exists(workbookPath))
        {
            Console.WriteLine($"Error: Workbook file not found – {workbookPath}");
            return;
        }

        if (!Directory.Exists(outputFolder))
        {
            try
            {
                Directory.CreateDirectory(outputFolder);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: Unable to create output folder – {ex.Message}");
                return;
            }
        }

        // ------------------------------------------------------------
        // 2. Load workbook
        // ------------------------------------------------------------
        Workbook workbook;
        try
        {
            workbook = new Workbook(workbookPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        // ------------------------------------------------------------
        // 3. Prepare HTML builder
        // ------------------------------------------------------------
        string htmlFilePath = Path.Combine(outputFolder, "WorkbookImageMap.html");
        using (StreamWriter htmlWriter = new StreamWriter(htmlFilePath, false))
        {
            htmlWriter.WriteLine("<!DOCTYPE html>");
            htmlWriter.WriteLine("<html lang=\"en\">");
            htmlWriter.WriteLine("<head>");
            htmlWriter.WriteLine("    <meta charset=\"UTF-8\">");
            htmlWriter.WriteLine("    <title>Workbook Image Map</title>");
            htmlWriter.WriteLine("</head>");
            htmlWriter.WriteLine("<body>");

            // ------------------------------------------------------------
            // 4. Process each worksheet
            // ------------------------------------------------------------
            for (int wsIndex = 0; wsIndex < workbook.Worksheets.Count; wsIndex++)
            {
                Worksheet sheet = workbook.Worksheets[wsIndex];
                string sheetNameSafe = MakeFileSystemSafe(sheet.Name);
                string tiffFileName = $"{sheetNameSafe}_page0.tiff";
                string tiffFullPath = Path.Combine(outputFolder, tiffFileName);

                // --------------------------------------------------------
                // 4.1 Render worksheet to a single‑page TIFF
                // --------------------------------------------------------
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    ImageType = Aspose.Cells.Drawing.ImageType.Tiff,
                    OnePagePerSheet = true,          // force one page per sheet
                    TiffCompression = TiffCompression.CompressionLZW,
                    HorizontalResolution = 150,
                    VerticalResolution = 150
                };

                SheetRender renderer = new SheetRender(sheet, imgOptions);
                try
                {
                    renderer.ToImage(0, tiffFullPath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to render sheet \"{sheet.Name}\": {ex.Message}");
                    continue;
                }

                // --------------------------------------------------------
                // 4.2 Build image tag with usemap attribute
                // --------------------------------------------------------
                string mapName = $"map_{wsIndex}";
                htmlWriter.WriteLine($"<h2>{System.Web.HttpUtility.HtmlEncode(sheet.Name)}</h2>");
                htmlWriter.WriteLine($"<img src=\"{tiffFileName}\" usemap=\"#{mapName}\" alt=\"{System.Web.HttpUtility.HtmlEncode(sheet.Name)}\"/>");

                // --------------------------------------------------------
                // 4.3 Build the <map> element based on comments
                // --------------------------------------------------------
                List<string> areaElements = new List<string>();

                foreach (Comment comment in sheet.Comments)
                {
                    // Only consider comments that have a non‑empty Note (treated as URL)
                    if (string.IsNullOrWhiteSpace(comment.Note))
                        continue;

                    // Determine the cell that the comment is anchored to
                    int row = comment.Row;
                    int col = comment.Column;

                    // Compute pixel coordinates of the cell rectangle
                    double left = GetPixelsBeforeColumn(sheet, col);
                    double top = GetPixelsBeforeRow(sheet, row);
                    double width = sheet.Cells.GetColumnWidthPixel(col);
                    double height = sheet.Cells.GetRowHeightPixel(row);

                    // Round to integer pixel values (required for HTML coords)
                    int leftInt = (int)Math.Round(left);
                    int topInt = (int)Math.Round(top);
                    int rightInt = (int)Math.Round(left + width);
                    int bottomInt = (int)Math.Round(top + height);

                    string url = comment.Note.Trim();
                    string alt = $"Cell {CellIndexToName(row, col)}";

                    string area = $"<area shape=\"rect\" coords=\"{leftInt},{topInt},{rightInt},{bottomInt}\" href=\"{System.Web.HttpUtility.HtmlAttributeEncode(url)}\" alt=\"{System.Web.HttpUtility.HtmlAttributeEncode(alt)}\" />";
                    areaElements.Add(area);
                }

                // Write the map only if we have at least one area
                if (areaElements.Count > 0)
                {
                    htmlWriter.WriteLine($"<map name=\"{mapName}\">");
                    foreach (string area in areaElements)
                    {
                        htmlWriter.WriteLine(area);
                    }
                    htmlWriter.WriteLine("</map>");
                }
                else
                {
                    htmlWriter.WriteLine($"<!-- No clickable comments found in sheet \"{sheet.Name}\" -->");
                }

                htmlWriter.WriteLine("<hr/>");
            }

            htmlWriter.WriteLine("</body>");
            htmlWriter.WriteLine("</html>");
        }

        Console.WriteLine($"HTML image map generated at: {htmlFilePath}");
    }

    // ------------------------------------------------------------
    // Helper: Compute total pixel width of all columns before the given index
    // ------------------------------------------------------------
    private static double GetPixelsBeforeColumn(Worksheet sheet, int columnIndex)
    {
        double total = 0.0;
        for (int c = 0; c < columnIndex; c++)
        {
            total += sheet.Cells.GetColumnWidthPixel(c);
        }
        return total;
    }

    // ------------------------------------------------------------
    // Helper: Compute total pixel height of all rows before the given index
    // ------------------------------------------------------------
    private static double GetPixelsBeforeRow(Worksheet sheet, int rowIndex)
    {
        double total = 0.0;
        for (int r = 0; r < rowIndex; r++)
        {
            total += sheet.Cells.GetRowHeightPixel(r);
        }
        return total;
    }

    // ------------------------------------------------------------
    // Helper: Convert zero‑based row/column to Excel style address (e.g., A1)
    // ------------------------------------------------------------
    private static string CellIndexToName(int row, int column)
    {
        return CellsHelper.CellIndexToName(row, column);
    }

    // ------------------------------------------------------------
    // Helper: Make a string safe for use as a file name (remove invalid chars)
    // ------------------------------------------------------------
    private static string MakeFileSystemSafe(string name)
    {
        foreach (char c in Path.GetInvalidFileNameChars())
        {
            name = name.Replace(c, '_');
        }
        return name;
    }
}
