// Title: Implement a custom IFilePathProvider for HtmlSaveOptions to preserve worksheet hyperlinks in HTML export (C#)
// Description: This example shows how to create a workbook with two sheets, add an internal hyperlink from Sheet1 to Sheet2, and assign a CustomFilePathProvider that returns "{SheetName}.html". By setting HtmlSaveOptions.FilePathProvider to this provider, each worksheet is saved as a separate HTML file and the hyperlink correctly points to the target sheet.
// Keywords: Aspose.Cells | IFilePathProvider | HtmlSaveOptions | HTML export | worksheet hyperlinks | C# | separate HTML files per sheet | internal link preservation
// Common Searches: Aspose.Cells keep internal hyperlinks when exporting to HTML | custom IFilePathProvider HtmlSaveOptions example | fix broken worksheet links after HTML export Aspose.Cells | save each worksheet as its own HTML file C# | HtmlSaveOptions FilePathProvider usage
// Developer Intent: Assign a custom IFilePathProvider to HtmlSaveOptions so that each worksheet is saved to its own HTML file and internal hyperlinks remain functional.
// Use Cases: Export a multi‑sheet workbook to separate HTML pages while retaining navigation links between sheets. | Generate web‑ready reports where users can click links to jump to related worksheets. | Publish Excel‑based dashboards as HTML with cross‑sheet references that stay active.
// AI Prompts: Provide C# code that implements a custom IFilePathProvider returning "{SheetName}.html" and applies it to HtmlSaveOptions. | Demonstrate how to verify that a hyperlink like "#Sheet2!A1" works after saving a workbook to HTML with a custom FilePathProvider. | Explain why internal hyperlinks break when a workbook is saved to a single HTML file and how a custom FilePathProvider resolves the issue.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Custom implementation of IFilePathProvider.
    // Returns a separate HTML file name for each worksheet,
    // ensuring that internal hyperlinks reference the correct files.
    // This example shows how to create a workbook with two sheets, add an internal hyperlink from Sheet1 to Sheet2, and assign a CustomFilePathProvider that returns "{SheetName}.html". By setting HtmlSaveOptions.FilePathProvider to this provider, each worksheet is saved as a separate HTML file and the hyperlink correctly points to the target sheet.
    public class CustomFilePathProvider : IFilePathProvider
    {
        public string GetFullName(string sheetName)
        {
            // Example: Sheet1 -> "Sheet1.html"
            return $"{sheetName}.html";
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook and add two worksheets.
            Workbook workbook = new Workbook();
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Sheet1";
            sheet1.Cells["A1"].PutValue("Go to Sheet2");

            // Add a hyperlink in Sheet1 that points to Sheet2!A1.
            // The link format "#SheetName!CellReference" works for internal links.
            sheet1.Hyperlinks.Add("A1", 1, 1, "#Sheet2!A1");

            // Create the second worksheet that will be the hyperlink target.
            Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
            sheet2.Cells["A1"].PutValue("Target Cell");

            // Configure HTML save options and assign the custom file path provider.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.FilePathProvider = new CustomFilePathProvider();

            // Save the workbook as HTML. Each worksheet will be saved to its own .html file,
            // and the hyperlinks will correctly reference those files.
            workbook.Save("WorkbookOutput.html", saveOptions);
        }
    }
}
