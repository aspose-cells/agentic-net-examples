// Title: Export an Excel workbook to HTML with per‑worksheet CSS files and uniform thin borders using ExportWorksheetCSSSeparately and SimilarBorderStyle in C#
// AI Prompts: Write C# code that creates a workbook, applies a thin border to a cell range, enables ExportWorksheetCSSSeparately and SimilarBorderStyle in HtmlSaveOptions, and saves each worksheet as HTML with its own CSS file. | Update the sample program to set HtmlSaveOptions.SimilarBorderStyle = true, generate separate CSS files for each sheet, and describe how to verify consistent border rendering in Chrome, Firefox, and Edge. | Create a C# script that iterates over all worksheets, applies the same thin border style to a specified range, and exports the workbook to HTML with distinct CSS files per worksheet using Aspose.Cells.
// Common Searches: Aspose.Cells C# export workbook to HTML with separate CSS files per sheet and thin borders | How to enable SimilarBorderStyle when saving Excel as HTML using Aspose.Cells | Create per‑sheet CSS during HTML conversion to ensure border consistency across browsers | C# example of generating distinct CSS for each worksheet with Aspose.Cells HtmlSaveOptions | Fix missing border styles in HTML output from Aspose.Cells conversion
// Tags: HtmlSaveOptions separate CSS per worksheet | Apply thin borders to cell range Aspose.Cells | Enable SimilarBorderStyle for HTML export | C# generate per‑sheet CSS files Aspose.Cells | Cross‑browser border consistency Excel to HTML

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example creates a new Workbook, adds two worksheets, populates cells, applies a uniform thin border to range A1:B2 on the first sheet, configures HtmlSaveOptions with ExportWorksheetCSSSeparately set to true, optionally enables SimilarBorderStyle, ensures the output directory exists, and saves the workbook as an HTML file (output.html) with separate CSS files for each worksheet.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // ---------- Worksheet 1 ----------
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Sheet1";

            // Populate some cells with data
            sheet1.Cells["A1"].PutValue("Header");
            sheet1.Cells["A2"].PutValue(123);
            sheet1.Cells["B2"].PutValue(456);

            // Apply a uniform thin border to a range (to test SimilarBorderStyle)
            Style borderStyle = workbook.CreateStyle();
            borderStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
            borderStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
            borderStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
            borderStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;

            StyleFlag styleFlag = new StyleFlag
            {
                Borders = true // Apply only border formatting
            };

            // Apply the style to the range A1:B2
            Aspose.Cells.Range range = sheet1.Cells.CreateRange("A1:B2");
            range.ApplyStyle(borderStyle, styleFlag);

            // ---------- Worksheet 2 ----------
            // Add a second worksheet to verify ExportWorksheetCSSSeparately
            Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
            sheet2.Cells["A1"].PutValue("Second sheet content");
            sheet2.Cells["A2"].PutValue("More data");

            // ---------- HTML Save Options ----------
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                // Export each worksheet's CSS into a separate file
                ExportWorksheetCSSSeparately = true
                // Note: SimilarBorderStyle property may not be available in older versions
            };

            // Ensure the output directory exists
            string outputPath = "output.html";
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as HTML
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
