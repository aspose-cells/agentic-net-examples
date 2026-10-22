// Title: Export each worksheet to a separate CSS file using HtmlSaveOptions.ExportWorksheetCSSSeparately in Aspose.Cells for .NET
// AI Prompts: Write C# code that saves a multi‑sheet workbook as HTML with HtmlSaveOptions.ExportWorksheetCSSSeparately set to true and prints the paths of the generated CSS files. | Show how to programmatically check for the existence of the per‑worksheet CSS files after exporting a workbook to HTML with Aspose.Cells. | Demonstrate creating a workbook with two worksheets, exporting it to HTML with separate CSS per sheet, and logging which CSS file corresponds to each worksheet.
// Common Searches: Aspose.Cells HtmlSaveOptions ExportWorksheetCSSSeparately example C# | how to generate separate CSS files for each worksheet when saving as HTML | verify per‑worksheet CSS files after HTML export Aspose.Cells | C# export multi‑sheet workbook to HTML with individual CSS files | Aspose.Cells separate CSS per sheet file naming convention
// Tags: HtmlSaveOptions ExportWorksheetCSSSeparately Aspose.Cells | per‑worksheet CSS files HTML export Aspose.Cells | verify worksheet CSS file existence Aspose.Cells | multi‑sheet workbook HTML export separate CSS | C# Aspose.Cells separate CSS per sheet

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample creates a workbook with two worksheets, enables HtmlSaveOptions.ExportWorksheetCSSSeparately, saves the workbook as HTML, and then confirms that CSS files named workbook.html_0.css and workbook.html_1.css exist for each worksheet.
class Program
{
    static void Main()
    {
        // Create a new workbook with two worksheets
        Workbook workbook = new Workbook();

        // First worksheet
        Worksheet sheet1 = workbook.Worksheets[0];
        sheet1.Name = "FirstSheet";
        sheet1.Cells["A1"].PutValue("Hello");

        // Second worksheet
        int sheet2Index = workbook.Worksheets.Add();
        Worksheet sheet2 = workbook.Worksheets[sheet2Index];
        sheet2.Name = "SecondSheet";
        sheet2.Cells["B2"].PutValue("World");

        // Configure HTML save options to export CSS separately for each worksheet
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
        htmlOptions.ExportWorksheetCSSSeparately = true;

        // Define output directory and ensure it exists
        string outputDir = "output";
        Directory.CreateDirectory(outputDir);

        // Save the workbook as HTML
        string htmlFilePath = Path.Combine(outputDir, "workbook.html");
        workbook.Save(htmlFilePath, htmlOptions);

        // Verify that each worksheet has its own CSS file
        // When ExportWorksheetCSSSeparately is true, Aspose.Cells creates files named:
        //   workbook.html_0.css, workbook.html_1.css, ... (index corresponds to worksheet index)
        foreach (Worksheet ws in workbook.Worksheets)
        {
            string cssFileName = $"workbook.html_{ws.Index}.css";
            string cssFilePath = Path.Combine(outputDir, cssFileName);

            if (File.Exists(cssFilePath))
            {
                Console.WriteLine($"CSS file for worksheet '{ws.Name}' exists: {cssFilePath}");
            }
            else
            {
                Console.WriteLine($"CSS file for worksheet '{ws.Name}' NOT found: {cssFilePath}");
            }
        }
    }
}
