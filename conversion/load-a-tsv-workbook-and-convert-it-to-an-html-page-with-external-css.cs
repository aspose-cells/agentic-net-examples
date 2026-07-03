using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the source TSV file
        string tsvFilePath = "input.tsv";

        // Load the TSV file into a workbook
        // LoadOptions with LoadFormat.Tsv ensures correct parsing of tab‑separated values
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Tsv);
        Workbook workbook = new Workbook(tsvFilePath, loadOptions);

        // Configure HTML save options to export worksheet CSS separately
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.ExportWorksheetCSSSeparately = true; // Generates an external CSS file

        // Optional: specify the folder where the external CSS file will be placed
        // By default it will be saved alongside the HTML file
        // htmlOptions.AttachedFilesDirectory = Path.GetDirectoryName(Path.GetFullPath("output.html"));

        // Save the workbook as an HTML page
        string htmlOutputPath = "output.html";
        workbook.Save(htmlOutputPath, htmlOptions);

        Console.WriteLine($"HTML page saved to: {Path.GetFullPath(htmlOutputPath)}");
        Console.WriteLine("External CSS file generated alongside the HTML page.");
    }
}