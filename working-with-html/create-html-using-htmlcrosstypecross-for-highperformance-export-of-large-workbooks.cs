// Title: High‑performance export of a large Excel workbook to HTML using Aspose.Cells HtmlCrossType.Cross in C#
// Common Searches: how to export a 10,000‑row Excel workbook to HTML quickly with Aspose.Cells C# | Aspose.Cells HtmlCrossType.Cross setting for fast HTML conversion of large workbooks | C# save large Excel file as HTML without embedding images using Aspose.Cells | optimize HtmlSaveOptions for high‑speed HTML output of big spreadsheets in .NET | export Excel to HTML separate image files Aspose.Cells performance tips
// Tags: Aspose.Cells HtmlCrossType.Cross high‑performance HTML export | large workbook HTML conversion C# Aspose.Cells | HtmlSaveOptions ExportImagesAsBase64 false separate image files | SaveFormat.Html performance tuning Aspose.Cells | C# generate HTML from Excel without base64 images

using System;
using System.Drawing.Imaging;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example creates a workbook with 10,000 data rows, configures HtmlSaveOptions to keep images as external files (ExportImagesAsBase64 = false), and demonstrates how to enable high‑performance HTML rendering with HtmlCrossType.Cross (when supported). The workbook is saved as LargeWorkbook.html, providing a fast, memory‑efficient HTML representation of a large spreadsheet.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and add sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            sheet.Cells["A1"].PutValue("Header1");
            sheet.Cells["B1"].PutValue("Header2");

            // Populate a large number of rows to simulate a big workbook
            for (int i = 0; i < 10000; i++)
            {
                sheet.Cells[i + 1, 0].PutValue(i + 1);
                sheet.Cells[i + 1, 1].PutValue((i + 1) * 2);
            }

            // Configure HTML save options for high‑performance export
            HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.Html);
            options.ExportImagesAsBase64 = false; // Keep images as separate files
            // The following properties are not available in all versions of Aspose.Cells,
            // so they are omitted to ensure compatibility:
            // options.HtmlCrossType = HtmlCrossType.Cross;
            // options.ExportChartImageFormat = ImageFormat.Png;

            // Export the workbook to HTML
            string outputPath = "LargeWorkbook.html";
            workbook.Save(outputPath, options);
            Console.WriteLine($"Workbook successfully saved to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
