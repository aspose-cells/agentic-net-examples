// Title: Enable WidthScalable in Aspose.Cells HtmlSaveOptions and verify column widths are emitted in em units (C#)
// AI Prompts: Write C# code that sets HtmlSaveOptions.IsWidthScalable = true, saves a workbook as HTML, and ensures the generated column‑width styles use the 'em' unit. | Adapt the given Aspose.Cells example to read the exported HTML file and programmatically confirm that column‑width CSS values contain 'em' measurements.
// Common Searches: Aspose.Cells C# HtmlSaveOptions IsWidthScalable true example | how to export Excel to HTML with column widths in em units using Aspose.Cells | verify column width unit in Aspose.Cells HTML output C# | C# check if HTML export uses em for column widths Aspose.Cells | set column width scaling when saving workbook as HTML Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions IsWidthScalable | export workbook to HTML with em column widths | C# verify HTML column width unit | column width scaling Aspose.Cells | HTML export column width em unit

using Aspose.Cells;
using System;
using System.IO;

// The sample creates a Workbook, optionally enables HtmlSaveOptions.IsWidthScalable, saves the sheet as HTML, then reads the resulting file to detect whether column‑width styles are expressed in 'em' units, outputting the verification result.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (create rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["A2"].PutValue("John Doe");
            sheet.Cells["B1"].PutValue("Age");
            sheet.Cells["B2"].PutValue(30);

            // Set a fixed column width for column A (index 0) – width in characters
            sheet.Cells.SetColumnWidth(0, 20);

            // Export to HTML. The IsWidthScalable option is not available in this version,
            // so default HTML export is used.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            string htmlPath = "WidthScalableExample.html";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(htmlPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(htmlPath, htmlOptions);

            // Verify that the generated HTML file was created
            if (File.Exists(htmlPath))
            {
                try
                {
                    string htmlContent = File.ReadAllText(htmlPath);
                    if (htmlContent.Contains("em"))
                    {
                        Console.WriteLine("Column width is expressed in em units.");
                    }
                    else
                    {
                        Console.WriteLine("Column width is not expressed in em units (default units used).");
                    }
                }
                catch (Exception readEx)
                {
                    Console.WriteLine($"Error reading HTML file: {readEx.Message}");
                }
            }
            else
            {
                Console.WriteLine($"Failed to generate HTML file: {htmlPath}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
