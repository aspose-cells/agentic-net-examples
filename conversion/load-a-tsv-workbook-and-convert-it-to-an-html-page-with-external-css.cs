// Title: Convert a TSV workbook to HTML with an external CSS file using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads a .tsv file with Aspose.Cells LoadOptions and saves it as an HTML document while exporting the worksheet CSS to a separate stylesheet. | Show how to configure HtmlSaveOptions to create an external CSS file and automatically create the output folder when converting a TSV workbook to HTML. | Write a console application that reads a tab‑separated values workbook and outputs both an HTML page and its accompanying CSS file using Aspose.Cells.
// Common Searches: asp.net convert tsv file to html with external css using aspose.cells | c# load tab separated values workbook and export html with separate stylesheet | htmlsaveoptions exportworksheetcssseparately example for tsv conversion | how to generate css file when saving workbook as html in aspose.cells c# | aspose.cells create output directory automatically when saving html
// Tags: TSV to HTML conversion Aspose.Cells | HtmlSaveOptions ExportWorksheetCSSSeparately | LoadOptions TSV format C# | Separate CSS generation Aspose.Cells | Create output directory on save

using System;
using System.IO;
using Aspose.Cells;

// // Loads a TSV workbook with LoadOptions, then saves it as HTML using HtmlSaveOptions with ExportWorksheetCSSSeparately=true, generating an external CSS file and creating the output folder if needed.
class TsvToHtmlWithExternalCss
{
    static void Main()
    {
        // Input TSV file path
        string inputTsvPath = "input.tsv";

        // Output HTML file path (CSS will be generated alongside this file)
        string outputHtmlPath = "output.html";

        // Load the TSV workbook using LoadOptions for TSV format
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Tsv);
        Workbook workbook = new Workbook(inputTsvPath, loadOptions);

        // Configure HTML save options to export worksheet CSS separately
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.ExportWorksheetCSSSeparately = true; // creates external CSS file
        htmlOptions.CreateDirectory = true; // auto‑create output folder if needed

        // Save the workbook as HTML; Aspose.Cells will generate an accompanying CSS file
        workbook.Save(outputHtmlPath, htmlOptions);

        Console.WriteLine("HTML file saved to: " + Path.GetFullPath(outputHtmlPath));
    }
}
