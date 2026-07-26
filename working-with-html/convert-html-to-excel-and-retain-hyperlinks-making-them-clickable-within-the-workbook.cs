// Title: C# – Convert HTML with Hyperlinks to Clickable Excel using Aspose.Cells
// Description: Shows how to load an HTML file containing anchor tags with Aspose.Cells (LoadOptions = LoadFormat.Html), automatically turn them into Excel Hyperlink objects, and save the workbook as XLSX so the links remain clickable.
// Keywords: Aspose.Cells | C# HTML to Excel | preserve hyperlinks | LoadOptions Html | clickable Excel links | convert HTML to XLSX | hyperlink conversion Aspose | automated report export
// Common Searches: Aspose.Cells keep hyperlinks when converting HTML to Excel | C# load HTML and export to XLSX with active links | How to preserve anchor tags in Excel using Aspose | LoadOptions Html hyperlink example | Convert web report HTML to Excel with clickable URLs
// Developer Intent: I need to transform an HTML document that contains links into an Excel workbook where the links stay functional.
// Use Cases: Turn a web‑generated report (HTML) into a downloadable Excel file while retaining the original URLs as active hyperlinks. | Automate conversion of email HTML content to Excel for data analysis, ensuring all embedded links remain operational. | Batch‑process a folder of HTML files, producing matching XLSX files that keep hyperlink interactivity.
// AI Prompts: Write C# code with Aspose.Cells that loads an HTML file containing hyperlinks and saves it as an XLSX workbook with clickable links. | Explain how LoadOptions set to LoadFormat.Html parses <a> tags into Excel Hyperlink objects in Aspose.Cells. | Provide a step‑by‑step guide to batch‑convert multiple HTML files to Excel while preserving hyperlink functionality.

using System;
using Aspose.Cells;

// Shows how to load an HTML file containing anchor tags with Aspose.Cells (LoadOptions = LoadFormat.Html), automatically turn them into Excel Hyperlink objects, and save the workbook as XLSX so the links remain clickable.
class HtmlToExcel
{
    static void Main()
    {
        // Path to the source HTML file containing hyperlinks
        string htmlFile = "input.html";

        // Path where the resulting Excel workbook will be saved
        string excelFile = "output.xlsx";

        // Load the HTML file into a Workbook.
        // LoadOptions with LoadFormat.Html ensures that hyperlinks in the HTML are parsed and
        // converted into Excel Hyperlink objects.
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Html);
        Workbook workbook = new Workbook(htmlFile, loadOptions);

        // Save the workbook in XLSX format. The hyperlinks are retained and become clickable
        // cells in the resulting Excel file.
        workbook.Save(excelFile, SaveFormat.Xlsx);
    }
}
