using System;
using System.IO;
using Aspose.Cells;

class ExportExcelToHtmlWithExternalImages
{
    static void Main()
    {
        // Path to the source Excel file
        string excelPath = "input.xlsx";

        // Load the workbook from the Excel file
        Workbook workbook = new Workbook(excelPath);

        // Configure HTML save options to export images as external files
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

        // Do not embed images as Base64; create separate image files instead
        htmlOptions.ExportImagesAsBase64 = false;

        // Directory where the exported image files will be saved
        htmlOptions.AttachedFilesDirectory = "output_images";

        // URL prefix used in the generated HTML to reference the image files
        htmlOptions.AttachedFilesUrlPrefix = "output_images/";

        // Ensure the image directory exists before saving
        Directory.CreateDirectory(htmlOptions.AttachedFilesDirectory);

        // Save the workbook as HTML using the configured options
        string htmlPath = "output.html";
        workbook.Save(htmlPath, htmlOptions);

        Console.WriteLine($"HTML file saved to '{htmlPath}'.");
        Console.WriteLine($"Background and other images are stored in '{htmlOptions.AttachedFilesDirectory}'.");
    }
}