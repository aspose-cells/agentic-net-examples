// Title: Generate separate CSS files for each worksheet when saving a workbook to HTML using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# save workbook as HTML with individual CSS files per worksheet | How to enable ExportWorksheetCSSSeparately in HtmlSaveOptions for .NET | Check existence of worksheet‑specific CSS files after HTML export with Aspose.Cells | C# code to create separate stylesheet files for each sheet using Aspose.Cells HtmlSaveOptions | Aspose.Cells HTML export separate stylesheet per sheet verification
// Tags: HtmlSaveOptions ExportWorksheetCSSSeparately option | Aspose.Cells separate CSS per worksheet | C# HTML export with individual stylesheet files | worksheet CSS file verification Aspose.Cells | generate per‑sheet CSS files .NET

using System;
using System.IO;
using Aspose.Cells;

// The example creates a workbook with two sheets, sets HtmlSaveOptions.ExportWorksheetCSSSeparately to true, saves the workbook as HTML, and then confirms that a distinct .css file (e.g., FirstSheet.css, SecondSheet.css) is generated for each worksheet in the output folder.
class Program
{
    static void Main()
    {
        // Create a new workbook with two worksheets
        Workbook workbook = new Workbook();
        workbook.Worksheets[0].Name = "FirstSheet";
        workbook.Worksheets[0].Cells["A1"].PutValue("Hello from first sheet");
        Worksheet secondSheet = workbook.Worksheets.Add("SecondSheet");
        secondSheet.Cells["A1"].PutValue("Hello from second sheet");

        // Configure HtmlSaveOptions to export CSS separately for each worksheet
        HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.Html);
        options.ExportWorksheetCSSSeparately = true;

        // Define output directory and ensure it exists
        string outputDir = "output";
        Directory.CreateDirectory(outputDir);

        // Save the workbook as HTML (main file)
        string htmlPath = Path.Combine(outputDir, "workbook.html");
        workbook.Save(htmlPath, options);

        // Verify that each worksheet has its own CSS file
        foreach (Worksheet ws in workbook.Worksheets)
        {
            // Expected CSS file name: {WorksheetName}.css
            string cssFileName = ws.Name + ".css";
            string cssFilePath = Path.Combine(outputDir, cssFileName);
            bool cssExists = File.Exists(cssFilePath);
            Console.WriteLine($"Worksheet '{ws.Name}' CSS file '{cssFileName}' exists: {cssExists}");
        }
    }
}
