// Title: Apply a corporate‑branded custom table style to the first ListObject in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an .xlsx file with Aspose.Cells, creates a Style with white bold font, corporate blue fill, thin black borders, and applies it to the DataRange of the first ListObject in the first worksheet. | Extend the sample to load style properties (font color, size, fill color, border style) from an external XML file and apply the generated style to every ListObject in the workbook.
// Common Searches: how to programmatically set a custom table style for a ListObject using Aspose.Cells C# | apply corporate branding colors to Excel tables with Aspose.Cells .NET example | load table style settings from XML and apply to all tables in a workbook using Aspose.Cells
// Tags: custom table formatting with Aspose.Cells C# | apply corporate style to ListObject data range | StyleFlag usage for table styling Aspose.Cells | read table style from XML Aspose.Cells | programmatic Excel table appearance .NET

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// Loads input.xlsx, defines a corporate‑branded Style (white bold font, blue fill, thin black borders), applies it to the first ListObject’s data range, and saves the workbook as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // -------------------------------------------------
            // Define a custom style that matches corporate branding.
            // -------------------------------------------------
            Style corporateStyle = workbook.CreateStyle();

            // Font settings
            corporateStyle.Font.Color = Color.White;
            corporateStyle.Font.Size = 12;
            corporateStyle.Font.IsBold = true;

            // Fill settings
            corporateStyle.ForegroundColor = Color.FromArgb(0, 112, 192); // corporate blue
            corporateStyle.Pattern = BackgroundType.Solid;

            // Border settings
            corporateStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
            corporateStyle.Borders[BorderType.BottomBorder].Color = Color.Black;
            corporateStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
            corporateStyle.Borders[BorderType.TopBorder].Color = Color.Black;
            corporateStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
            corporateStyle.Borders[BorderType.LeftBorder].Color = Color.Black;
            corporateStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
            corporateStyle.Borders[BorderType.RightBorder].Color = Color.Black;

            // -------------------------------------------------
            // Apply the custom style to the first table (ListObject) in the first worksheet.
            // -------------------------------------------------
            Worksheet sheet = workbook.Worksheets[0];

            if (sheet.ListObjects.Count > 0)
            {
                // Get the first table
                ListObject table = sheet.ListObjects[0];

                // Optionally set a built‑in table style (skip if custom style is not a built‑in one)
                // table.TableStyleName = "TableStyleMedium2";

                // Apply the custom style to the data range of the table for visual consistency
                StyleFlag flag = new StyleFlag { All = true };
                table.DataRange.ApplyStyle(corporateStyle, flag);
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
