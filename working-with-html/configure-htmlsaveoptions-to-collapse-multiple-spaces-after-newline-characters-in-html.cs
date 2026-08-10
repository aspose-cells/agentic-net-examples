// Title: Collapse extra spaces after line breaks when exporting Excel to HTML with Aspise.Cells for .NET
// Description: Shows how to save a workbook as HTML using HtmlSaveOptions and then eliminate redundant spaces after newline (<br>) tags by loading the file with HtmlLoadOptions.DeleteRedundantSpaces in C#.
// Keywords: Aspose.Cells | HtmlSaveOptions | collapse spaces | DeleteRedundantSpaces | C# | export Excel to HTML | remove extra spaces | .NET | line break spaces
// Common Searches: Aspose.Cells collapse spaces after newline | HtmlLoadOptions DeleteRedundantSpaces C# example | remove extra spaces in HTML exported from Excel | how to trim spaces after <br> tag using Aspose.Cells | export active worksheet to HTML without redundant spaces
// Developer Intent: Generate HTML from a workbook where multiple spaces following line‑break characters are automatically collapsed.
// Use Cases: Export only the active worksheet to HTML while ensuring whitespace after <br> tags is trimmed. | Reload the saved HTML with DeleteRedundantSpaces to obtain clean cell text for further processing. | Automate HTML export on a server and guarantee consistent spacing across different locales.
// AI Prompts: Provide a C# snippet that saves an Aspose.Cells workbook to HTML and collapses extra spaces after line breaks. | Explain how HtmlLoadOptions.DeleteRedundantSpaces works as a workaround for missing space‑collapse settings in HtmlSaveOptions. | Show step‑by‑step code to export a worksheet to HTML, then read it back with redundant spaces removed.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlSpaceCollapseDemo
{
    // Shows how to save a workbook as HTML using HtmlSaveOptions and then eliminate redundant spaces after newline (<br>) tags by loading the file with HtmlLoadOptions.DeleteRedundantSpaces in C#.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Put a string that contains a newline and multiple spaces after it
            // Example: "Line1\n   Line2"
            sheet.Cells["A1"].PutValue("First line\n   Second line with   extra spaces");

            // Configure HTML save options
            // Aspose.Cells does not provide a direct property to collapse spaces after newline during saving.
            // However, we can set general options here. The example sets ExportActiveWorksheetOnly for brevity.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.ExportActiveWorksheetOnly = true;

            // Save the workbook as HTML
            string htmlPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Output.html");
            workbook.Save(htmlPath, saveOptions);

            // To demonstrate collapsing redundant spaces after line breaks, load the HTML with HtmlLoadOptions.
            // This step is optional and shows how to remove extra spaces when reading the HTML back.
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();
            loadOptions.DeleteRedundantSpaces = true; // Collapse multiple spaces after <br> tags

            // Load the saved HTML back into a new workbook
            Workbook loadedWorkbook = new Workbook(htmlPath, loadOptions);
            string cellValue = loadedWorkbook.Worksheets[0].Cells["A1"].StringValue;

            Console.WriteLine("Cell value after loading with DeleteRedundantSpaces:");
            Console.WriteLine(cellValue);
        }
    }
}
