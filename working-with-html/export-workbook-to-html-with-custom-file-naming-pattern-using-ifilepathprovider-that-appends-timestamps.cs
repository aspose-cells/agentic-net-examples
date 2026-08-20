// Title: Export Aspose.Cells Workbook to HTML with Timestamped Worksheet Filenames via IFilePathProvider (C#)
// Description: Demonstrates a custom TimestampFilePathProvider that implements IFilePathProvider, appends a yyyyMMdd_HHmmss suffix to each worksheet name, and assigns it to HtmlSaveOptions.FilePathProvider so that workbook.Save creates separate HTML files with unique timestamped names.
// Keywords: Aspose.Cells | HTML export | IFilePathProvider | timestamp filename | custom file naming | C# | .NET | HtmlSaveOptions | separate worksheet files | date‑time suffix
// Common Searches: Aspose.Cells IFilePathProvider example | C# export workbook to HTML with custom file names | Add timestamp to HTML files when saving Aspose.Cells workbook | Save each worksheet as separate HTML file Aspose.Cells | HtmlSaveOptions custom file path provider C#
// Developer Intent: Export a workbook to HTML where each worksheet is saved as an individual file using a timestamp‑based naming pattern.
// Use Cases: Automated daily reporting that creates uniquely timestamped HTML snapshots for each sheet. | Archiving multi‑sheet financial models as separate HTML files to support version control and audit trails. | Generating web‑ready worksheet pages without overwriting previous exports by embedding the generation time in the filename.
// AI Prompts: Write C# code that uses IFilePathProvider to add a custom prefix and a UTC timestamp to worksheet HTML filenames in Aspose.Cells. | Explain how HtmlSaveOptions.FilePathProvider influences the creation of separate HTML files during workbook.Save. | Show how to modify TimestampFilePathProvider to use a different date format, such as yyyy-MM-dd_HH-mm-ss, and to store files in a specific folder.

using System;
using Aspose.Cells;

namespace AsposeCellsExportHtml
{
    // Custom file path provider that appends a timestamp to each worksheet file name.
    // Demonstrates a custom TimestampFilePathProvider that implements IFilePathProvider, appends a yyyyMMdd_HHmmss suffix to each worksheet name, and assigns it to HtmlSaveOptions.FilePathProvider so that workbook.Save creates separate HTML files with unique timestamped names.
    public class TimestampFilePathProvider : IFilePathProvider
    {
        public string GetFullName(string sheetName)
        {
            // Create a timestamp string (e.g., 20230811_153045).
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            // Return the file name with the timestamp and .html extension.
            return $"{sheetName}_{timestamp}.html";
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();

            // Populate data in the first worksheet.
            Worksheet ws1 = workbook.Worksheets[0];
            ws1.Name = "Report";
            ws1.Cells["A1"].PutValue("Item");
            ws1.Cells["B1"].PutValue("Quantity");
            ws1.Cells["A2"].PutValue("Apples");
            ws1.Cells["B2"].PutValue(120);
            ws1.Cells["A3"].PutValue("Oranges");
            ws1.Cells["B3"].PutValue(85);

            // Add a second worksheet to demonstrate separate file naming.
            Worksheet ws2 = workbook.Worksheets.Add("Summary");
            ws2.Cells["A1"].PutValue("Total Items");
            ws2.Cells["B1"].PutValue(205);

            // Configure HTML save options and assign the custom file path provider.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.FilePathProvider = new TimestampFilePathProvider();

            // Save the workbook as HTML. Each worksheet will be saved using the custom naming pattern.
            workbook.Save("Workbook.html", saveOptions);
        }
    }
}
