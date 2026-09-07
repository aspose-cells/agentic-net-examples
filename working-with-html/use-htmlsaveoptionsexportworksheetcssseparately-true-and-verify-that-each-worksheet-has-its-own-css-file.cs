// Title: Save an Aspose.Cells workbook as HTML with a separate CSS file for each worksheet using C#
// AI Prompts: Generate C# code that creates a workbook with multiple sheets and saves it to HTML using HtmlSaveOptions with ExportWorksheetCSSSeparately set to true, producing a distinct .css file for each worksheet. | Write a C# snippet that scans the output folder after an Aspose.Cells HTML export and confirms that a .css file named after each worksheet exists. | Show how to adapt an existing Aspose.Cells HTML export routine to enable per‑worksheet CSS files and direct all output to a specified directory.
// Common Searches: Aspose.Cells C# export workbook to html with separate css per worksheet example | how to enable ExportWorksheetCSSSeparately in Aspose.Cells and verify css files | C# Aspose.Cells HTML save options per‑worksheet css generation | checking css files for each sheet after Aspose.Cells HTML conversion
// Tags: HtmlSaveOptions ExportWorksheetCSSSeparately C# | per‑worksheet css generation Aspose.Cells | save workbook as html with individual css files | verify generated css files Aspose.Cells | multiple worksheets html export Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The program creates a workbook with two worksheets, configures HtmlSaveOptions to export each sheet's styling to its own .css file, saves the workbook as HTML, and then checks the output directory to ensure a CSS file exists for every worksheet.
class Program
{
    static void Main()
    {
        // Create a new workbook with two worksheets
        Workbook workbook = new Workbook();
        Worksheet sheet1 = workbook.Worksheets[0];
        sheet1.Name = "FirstSheet";
        sheet1.Cells["A1"].PutValue("Hello from Sheet 1");

        Worksheet sheet2 = workbook.Worksheets.Add("SecondSheet");
        sheet2.Cells["A1"].PutValue("Hello from Sheet 2");

        // Configure HTML save options to export CSS separately for each worksheet
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
        htmlOptions.ExportWorksheetCSSSeparately = true; // each worksheet gets its own .css file
        htmlOptions.ExportActiveWorksheetOnly = false;   // export all worksheets

        // Define output folder and base file name
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "HtmlOutput");
        Directory.CreateDirectory(outputFolder);
        string htmlFilePath = Path.Combine(outputFolder, "Workbook.html");

        // Save the workbook as HTML
        workbook.Save(htmlFilePath, htmlOptions);

        // Verify that each worksheet has its own CSS file
        foreach (Worksheet ws in workbook.Worksheets)
        {
            // Aspose.Cells names the CSS file as "<WorksheetName>.css"
            string cssFileName = ws.Name + ".css";
            string cssFilePath = Path.Combine(outputFolder, cssFileName);

            if (File.Exists(cssFilePath))
            {
                Console.WriteLine($"CSS file found for worksheet \"{ws.Name}\": {cssFileName}");
            }
            else
            {
                Console.WriteLine($"CSS file NOT found for worksheet \"{ws.Name}\". Expected: {cssFileName}");
            }
        }
    }
}
