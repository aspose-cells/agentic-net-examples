// Title: Merge Cells D8:E10, Apply Fraction Format, and Export as PNG with Aspose.Cells for .NET
// Description: Loads an existing workbook, merges the range D8:E10, applies a custom fraction number format ("# ?/?"), renders the first worksheet page to a PNG image, and optionally saves the modified file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# | merge cells D8:E10 | fraction number format | custom number format Excel | export worksheet to PNG | SheetRender PNG | ImageOrPrintOptions | .NET Excel to image | render Excel as image | Aspose.Cells rendering options
// Common Searches: how to merge cells and export to png with aspose.cells | apply fraction format to merged cells in c# | aspose.cells render worksheet as png image | export excel range as png using asp.net | c# aspose.cells custom number format image export
// Developer Intent: Merge D8:E10, set a fraction format, and generate a PNG snapshot of the worksheet using Aspose.Cells for .NET.
// Use Cases: Create printable reports where merged cells show fractional values and the sheet is delivered as an image. | Generate thumbnail previews of formatted Excel data for web dashboards or mobile apps. | Automate conversion of Excel worksheets with custom number formats into PNG files for email or document attachments.
// AI Prompts: Show how to export only the merged range D8:E10 as a PNG instead of the whole sheet. | Provide code to set image resolution, DPI, and background color when rendering with Aspose.Cells. | Explain how to apply the fraction style to a named range and then render that range to a PNG image.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Loads an existing workbook, merges the range D8:E10, applies a custom fraction number format ("# ?/?"), renders the first worksheet page to a PNG image, and optionally saves the modified file using Aspose.Cells for .NET.
class MergeFractionAndExportImage
{
    static void Main()
    {
        try
        {
            // Input and output file paths
            string inputFile = "input.xlsx";   // existing workbook to load
            string outputImage = "merged_fraction.png";

            // Verify input file exists
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Input file \"{inputFile}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputFile);
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Merge cells D8:E10 (rows 7‑9, columns 3‑4)
            cells.Merge(7, 3, 3, 2);

            // Create a fraction number format "# ?/?" and apply it to the merged range
            Style fractionStyle = workbook.CreateStyle();
            fractionStyle.Custom = "# ?/?"; // fraction format
            StyleFlag flag = new StyleFlag { NumberFormat = true };

            // Apply the style to each cell in the merged area
            for (int row = 7; row < 7 + 3; row++)
            {
                for (int col = 3; col < 3 + 2; col++)
                {
                    cells[row, col].SetStyle(fractionStyle, flag);
                }
            }

            // Configure image export options
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();
            // Image format is inferred from the file extension; explicit setting omitted for compatibility
            imgOptions.OnePagePerSheet = true;

            // Render the worksheet (first page) to a PNG image
            SheetRender renderer = new SheetRender(worksheet, imgOptions);
            renderer.ToImage(0, outputImage); // export page 0 to PNG file

            // Optionally save the modified workbook
            workbook.Save("modified.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
