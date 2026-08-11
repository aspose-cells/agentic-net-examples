// Title: Custom IFilePathProvider for Aspose.Cells HTML Export – Separate Sheet Files & Preserved Links
// Description: Demonstrates how to implement IFilePathProvider to generate custom relative paths (e.g., "Sheets/Sheet1.html") for each worksheet when saving a workbook to HTML with Aspose.Cells for .NET. The example configures HtmlSaveOptions to use the provider, export all worksheets, and create full‑path links so inter‑sheet hyperlinks remain functional.
// Keywords: Aspose.Cells IFilePathProvider | HTML export custom folder | C# Aspose.Cells separate sheet files | preserve inter‑sheet hyperlinks | HtmlSaveOptions IsFullPathLink | export all worksheets to HTML | custom file path provider example | Aspose.Cells HTML link integrity
// Common Searches: Aspose.Cells custom HTML file path per worksheet | How to keep hyperlinks when exporting workbook to HTML | IFilePathProvider example C# | Save each Excel sheet as separate HTML file Aspose | Full path links Aspose.Cells HTML export
// Developer Intent: Create a custom IFilePathProvider to control where each worksheet’s HTML file is saved and ensure hyperlink references stay valid.
// Use Cases: Publish a multi‑sheet workbook as independent HTML pages stored in a dedicated subfolder while maintaining clickable links between sheets. | Automate documentation pipelines that require predictable file names (e.g., SheetName.html) for each worksheet. | Move the main HTML file to a different location without breaking inter‑sheet links by enabling full‑path linking.
// AI Prompts: Generate C# code that implements IFilePathProvider to save worksheets as "Sheets/<WorksheetName>.html" and configures HtmlSaveOptions for full‑path links. | Show how to export an Aspose.Cells workbook to HTML with ExportActiveWorksheetOnly = false and a custom file path provider, preserving all inter‑sheet hyperlinks. | Explain how to modify the custom file path provider to include a date‑stamp folder (e.g., "Sheets/2024-08-11/Sheet1.html").

using System;
using Aspose.Cells;

namespace AsposeCellsCustomPathDemo
{
    // Custom implementation of IFilePathProvider.
    // Generates a relative path for each worksheet HTML file.
    // Demonstrates how to implement IFilePathProvider to generate custom relative paths (e.g., "Sheets/Sheet1.html") for each worksheet when saving a workbook to HTML with Aspose.Cells for .NET. The example configures HtmlSaveOptions to use the provider, export all worksheets, and create full‑path links so inter‑sheet hyperlinks remain functional.
    public class CustomFilePathProvider : IFilePathProvider
    {
        // sheetName – name of the worksheet being exported.
        // Returns a path like "Sheets/Sheet1.html".
        public string GetFullName(string sheetName)
        {
            // You can customize the folder or naming scheme as needed.
            return $"Sheets/{sheetName}.html";
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a workbook with three worksheets.
            Workbook wb = new Workbook();
            wb.Worksheets[0].Name = "Summary";
            wb.Worksheets.Add("Data");
            wb.Worksheets.Add("Report");

            // Populate each sheet with sample data.
            wb.Worksheets["Summary"].Cells["A1"].PutValue("This is the summary sheet.");
            wb.Worksheets["Data"].Cells["A1"].PutValue("Data sheet content.");
            wb.Worksheets["Report"].Cells["A1"].PutValue("Report sheet content.");

            // Set up HTML save options.
            HtmlSaveOptions options = new HtmlSaveOptions();
            // Use the custom file path provider so each worksheet is saved to its own file.
            options.FilePathProvider = new CustomFilePathProvider();
            // Export all worksheets (not only the active one) to keep inter‑sheet links.
            options.ExportActiveWorksheetOnly = false;
            // Use full path links to ensure references remain valid regardless of location.
            options.IsFullPathLink = true;

            // Save the workbook to HTML. The main file will be "Workbook.html",
            // and each worksheet will be saved to the paths returned by the provider.
            wb.Save("Workbook.html", options);
        }
    }
}
