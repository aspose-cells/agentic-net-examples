// Title: C# – Custom IFilePathProvider for lowercase HTML filenames per worksheet (Aspose.Cells)
// Description: Demonstrates how to implement a LowerCaseFilePathProvider that sanitizes worksheet names with CellsHelper.CreateSafeSheetName, converts them to lower‑case, adds a .html extension, and plugs the provider into HtmlSaveOptions so each sheet is saved as a URL‑friendly HTML file.
// Keywords: Aspose.Cells | C# | IFilePathProvider | HtmlSaveOptions | lowercase filenames | URL‑friendly file names | CreateSafeSheetName | HTML export per worksheet | custom file path provider | Aspose.Cells example
// Common Searches: Aspose.Cells custom IFilePathProvider example | how to generate lowercase HTML files for worksheets | C# export each worksheet to separate HTML file Aspose.Cells | make worksheet HTML filenames URL safe | HtmlSaveOptions file path provider C#
// Developer Intent: Create a custom IFilePathProvider that turns worksheet names into safe, lower‑case .html file names for consistent linking during HTML export.
// Use Cases: Export a multi‑sheet workbook where each sheet is saved as a predictable, lowercase HTML file for SEO‑friendly URLs. | Ensure generated file names are safe for file systems and web servers by using CellsHelper.CreateSafeSheetName and ToLowerInvariant. | Integrate the provider into HtmlSaveOptions so the naming convention is applied automatically on workbook.Save.
// AI Prompts: Write a C# class implementing Aspose.Cells IFilePathProvider that returns safe, lower‑case .html filenames based on worksheet names. | Show how to assign a custom IFilePathProvider to HtmlSaveOptions and export a workbook so each worksheet is saved as a separate lowercase HTML file. | Explain the role of CellsHelper.CreateSafeSheetName and ToLowerInvariant in producing URL‑compatible file names for HTML export with Aspose.Cells.

using Aspose.Cells;
using System;

namespace AsposeCellsFilePathProviderDemo
{
    // Custom implementation of IFilePathProvider that creates lowercase HTML file names
    // Demonstrates how to implement a LowerCaseFilePathProvider that sanitizes worksheet names with CellsHelper.CreateSafeSheetName, converts them to lower‑case, adds a .html extension, and plugs the provider into HtmlSaveOptions so each sheet is saved as a URL‑friendly HTML file.
    internal class LowerCaseFilePathProvider : IFilePathProvider
    {
        public string GetFullName(string sheetName)
        {
            // Ensure the sheet name is a valid file name
            string safeName = CellsHelper.CreateSafeSheetName(sheetName);
            // Convert to lower case and append .html extension
            return $"{safeName.ToLowerInvariant()}.html";
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and add some worksheets with various names
            Workbook workbook = new Workbook();
            workbook.Worksheets[0].Name = "Summary";
            workbook.Worksheets.Add("Data Sheet");
            workbook.Worksheets.Add("Charts&Analysis");

            // Optional: put some data into the worksheets
            workbook.Worksheets["Summary"].Cells["A1"].PutValue("Summary content");
            workbook.Worksheets["Data Sheet"].Cells["A1"].PutValue("Data content");
            workbook.Worksheets["Charts&Analysis"].Cells["A1"].PutValue("Chart content");

            // Set up HTML save options to use the custom file path provider
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            // Export each worksheet to a separate HTML file
            saveOptions.ExportActiveWorksheetOnly = false;
            // Assign the custom provider
            saveOptions.FilePathProvider = new LowerCaseFilePathProvider();

            // Save the workbook; each worksheet will be saved as a lowercase .html file
            workbook.Save("WorkbookOutput.html", saveOptions);
        }
    }
}
