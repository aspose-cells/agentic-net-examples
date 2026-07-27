// Title: C# – Set Arial as Default Font for HTML Export with Aspose.Cells HtmlSaveOptions
// Description: Demonstrates how to create a workbook, add sample data, configure HtmlSaveOptions.DefaultFontName to "Arial", and save the workbook as an HTML file using Aspose.Cells for .NET. The example ensures all generated HTML text uses Arial without altering individual cell styles.
// Keywords: Aspose.Cells | HtmlSaveOptions | DefaultFontName | Arial font | HTML export C# | Excel to HTML conversion | Aspose.Cells example | C# workbook to HTML | set default font Aspose | Aspose.Cells HTML output
// Common Searches: Aspose.Cells set default font for HTML export | HtmlSaveOptions.DefaultFontName C# example | Export Excel to HTML with Arial using Aspose | C# code to force Arial font in HTML output | Aspose.Cells HTML conversion default font setting
// Developer Intent: Configure Aspose.Cells to use Arial as the default font when saving a workbook to HTML.
// Use Cases: Produce web‑ready reports from Excel files with a consistent Arial typeface for branding. | Automate batch conversion of multiple workbooks to HTML while enforcing a single default font. | Create lightweight HTML previews of spreadsheets for email or intranet publishing without editing each cell’s style.
// AI Prompts: Generate C# code that exports an Aspose.Cells workbook to HTML with the default font set to Arial and adds custom CSS for table borders. | Explain the effect of HtmlSaveOptions.DefaultFontName on the generated HTML and how to override it for specific cells or ranges. | Show how to change the default font to a different type (e.g., "Calibri") while preserving existing cell formatting during HTML export.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, add sample data, configure HtmlSaveOptions.DefaultFontName to "Arial", and save the workbook as an HTML file using Aspose.Cells for .NET. The example ensures all generated HTML text uses Arial without altering individual cell styles.
    public class SetDefaultFontHtmlExport
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add sample data to demonstrate the font effect
                worksheet.Cells["A1"].PutValue("Sample Text for HTML Export");

                // Set HTML save options with default font Arial
                HtmlSaveOptions saveOptions = new HtmlSaveOptions();
                saveOptions.DefaultFontName = "Arial";

                string outputPath = "output.html";

                // Save the workbook as HTML
                workbook.Save(outputPath, saveOptions);
                Console.WriteLine($"Workbook saved successfully to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            SetDefaultFontHtmlExport.Run();
        }
    }
}
