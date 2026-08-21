// Title: Use a Custom IFilePathProvider with HtmlSaveOptions to Preserve Cross‑Sheet Hyperlinks in HTML Export (C#)
// Description: Demonstrates how to implement a custom IFilePathProvider, assign it to HtmlSaveOptions.FilePathProvider, and export each worksheet to its own HTML file while keeping hyperlinks between sheets functional.
// Keywords: Aspose.Cells | IFilePathProvider | HtmlSaveOptions | custom file path provider | cross‑sheet hyperlink | HTML export | separate worksheet files | C# | .NET | web reporting
// Common Searches: Aspose.Cells custom IFilePathProvider example | fix broken hyperlinks when exporting to HTML | save each worksheet as separate HTML file Aspose | HtmlSaveOptions FilePathProvider usage | preserve cross sheet links in HTML export
// Developer Intent: Assign a custom IFilePathProvider to HtmlSaveOptions so that hyperlinks between worksheets point to the correct HTML files after each sheet is saved separately.
// Use Cases: Generate HTML reports where every worksheet resides in its own file within a specific folder structure. | Maintain navigation links between sheets in a multi‑page web view of a workbook. | Integrate Aspose.Cells HTML export into web applications that require custom naming or placement of worksheet files.
// AI Prompts: Show me how to create an IFilePathProvider that returns a subfolder path for each worksheet when saving to HTML with Aspose.Cells. | Provide a complete C# example that uses HtmlSaveOptions.FilePathProvider to keep cross‑sheet hyperlinks working after export. | Explain how to modify the custom file path provider to use absolute paths or different file extensions for the generated HTML files.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Custom implementation of IFilePathProvider.
    // Returns a full file name for each worksheet when exporting to HTML separately.
    // Demonstrates how to implement a custom IFilePathProvider, assign it to HtmlSaveOptions.FilePathProvider, and export each worksheet to its own HTML file while keeping hyperlinks between sheets functional.
    public class CustomFilePathProvider : IFilePathProvider
    {
        public string GetFullName(string sheetName)
        {
            // Example: place each worksheet HTML file in a "sheets" subfolder.
            // Adjust the path as needed for your environment.
            return $"sheets\\{sheetName}.html";
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook and access the first worksheet.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "FirstSheet";

            // Add some sample data.
            sheet.Cells["A1"].PutValue("Hello");
            sheet.Cells["A2"].PutValue("World");

            // Add a hyperlink that points to the second worksheet (which will be exported separately).
            // The hyperlink will be fixed by the custom IFilePathProvider.
            sheet.Hyperlinks.Add("B1", 1, 1, "SecondSheet!A1");

            // Add a second worksheet to demonstrate cross‑sheet linking.
            Worksheet secondSheet = workbook.Worksheets.Add("SecondSheet");
            secondSheet.Cells["A1"].PutValue("Target Cell");

            // Configure HTML save options and assign the custom file path provider.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.FilePathProvider = new CustomFilePathProvider();

            // Save the workbook as HTML. Each worksheet will be saved to its own file
            // using the paths returned by CustomFilePathProvider.
            workbook.Save("output.html", saveOptions);

            Console.WriteLine("Workbook saved to HTML with custom file paths.");
        }
    }
}
