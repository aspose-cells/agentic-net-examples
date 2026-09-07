// Title: Export a specific print area to HTML with CSS disabled using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that sets a worksheet's PrintArea to a range and saves the workbook as HTML with HtmlSaveOptions.DisableCss enabled. | Show how to configure Aspose.Cells HtmlSaveOptions to prevent CSS file creation while exporting only the defined print area. | Provide a minimal C# example that creates a workbook, defines a print area, disables CSS generation, and writes the output to an HTML file.
// Common Searches: how to export only a selected range to html using aspose.cells c# | asp.net disable css generation when saving workbook as html with aspose.cells | set print area before html export in aspose.cells .net | htmlsaveoptions.disablecss example for c# | export workbook to html without external stylesheet asp.net
// Tags: Aspose.Cells HtmlSaveOptions.DisableCss | export worksheet print area to HTML | C# Aspose.Cells HTML export without CSS | define print area Aspose.Cells | save workbook as HTML Aspose.Cells .NET

using Aspose.Cells;

// // This program creates a new workbook, defines a print area (A1:B2), configures HtmlSaveOptions to disable CSS generation, and saves the workbook as an HTML file.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // (Optional) Populate some data
        sheet.Cells["A1"].PutValue("Header");
        sheet.Cells["A2"].PutValue(123);
        sheet.Cells["B2"].PutValue(456);

        // Define the print area (e.g., cells A1:B2)
        sheet.PageSetup.PrintArea = "A1:B2";

        // Configure HTML save options to disable CSS generation
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.DisableCss = true; // No external CSS file will be created

        // Export the workbook to HTML
        workbook.Save("output.html", htmlOptions);
    }
}
