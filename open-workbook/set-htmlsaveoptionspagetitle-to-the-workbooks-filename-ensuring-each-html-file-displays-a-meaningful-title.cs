// Title: Set HtmlSaveOptions.PageTitle to the workbook filename when exporting Excel to HTML with Aspose.Cells for .NET
// AI Prompts: Generate C# code that opens an Excel file using Aspose.Cells, assigns HtmlSaveOptions.PageTitle to the workbook's file name, and saves the workbook as an HTML document. | Write C# that verifies the existence of the source Excel file, creates a simple workbook if it is missing, then exports it to HTML with the page title derived from the source file name. | Show how to construct the output HTML file path by replacing the source workbook’s extension with .html and pass it to Workbook.Save with the configured HtmlSaveOptions.
// Common Searches: aspnet set html page title from excel filename using aspose.cells | c# export workbook to html with dynamic title based on file name | how to use HtmlSaveOptions.PageTitle property in Aspose.Cells .NET | fallback create workbook when source excel missing then save as html Aspose.Cells | change excel file extension to html path c# Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions.PageTitle | export Excel to HTML C# | dynamic HTML title from workbook filename | fallback workbook creation Aspose.Cells | generate HTML file path from Excel path C#

using System;
using System.IO;
using Aspose.Cells;

// C# example that loads an existing Excel file (or creates a minimal workbook if the file is absent), sets HtmlSaveOptions.PageTitle to the workbook's filename, and saves the workbook as an HTML file with that title.
class HtmlExportWithTitle
{
    static void Main()
    {
        // Path to the source Excel file
        string excelPath = @"C:\Data\Report.xlsx";

        try
        {
            Workbook workbook;

            // Ensure the source file exists; if not, create a simple workbook
            if (File.Exists(excelPath))
            {
                workbook = new Workbook(excelPath);
            }
            else
            {
                // Create a new workbook with a sample sheet
                workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "SampleSheet";
                sheet.Cells["A1"].PutValue("This is a generated workbook because the source file was not found.");
                // Save the generated workbook so subsequent runs can use it
                Directory.CreateDirectory(Path.GetDirectoryName(excelPath));
                workbook.Save(excelPath);
            }

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Set the page title to the workbook's filename (including extension)
                PageTitle = Path.GetFileName(excelPath)
            };

            // Define the output HTML file path
            string htmlPath = Path.ChangeExtension(excelPath, ".html");

            // Save the workbook as HTML using the configured options
            workbook.Save(htmlPath, htmlOptions);

            Console.WriteLine($"Workbook successfully exported to HTML: {htmlPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during export: {ex.Message}");
        }
    }
}
