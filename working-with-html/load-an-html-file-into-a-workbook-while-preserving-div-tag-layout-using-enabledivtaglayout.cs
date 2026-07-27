// Title: Load HTML with DIV Layout into Aspose.Cells Workbook (C#) using HtmlLoadOptions.SupportDivTag
// Description: Shows how to import an HTML file into an Aspose.Cells Workbook while preserving the original <div> layout by enabling SupportDivTag in HtmlLoadOptions, then export the result to XLSX.
// Keywords: Aspose.Cells | HtmlLoadOptions | SupportDivTag | load HTML to Excel | preserve DIV layout | C# HTML to XLSX conversion | HTML import Aspose.Cells | div tag layout Excel | .NET Excel export HTML
// Common Searches: Aspose.Cells load HTML with div layout C# | HtmlLoadOptions SupportDivTag example | preserve <div> structure when converting HTML to Excel | how to keep div positioning in Aspose.Cells import | convert web page to XLSX without losing div formatting
// Developer Intent: Import an HTML document into a workbook while retaining the visual arrangement created by <div> elements.
// Use Cases: Transform a web‑based report that uses positioned DIVs into an Excel file for offline analysis without altering its layout. | Batch‑process a collection of HTML dashboards that rely on DIV positioning, generating XLSX files that mirror the original design. | Validate that the content of the top‑left DIV in the source HTML appears in cell A1 after conversion.
// AI Prompts: Generate a C# snippet that loads an HTML file with SupportDivTag enabled and then auto‑fits all columns. | Provide code to import HTML using HtmlLoadOptions, log any conversion warnings, and save the workbook as XLSX. | Explain how to combine SupportDivTag with CSS style preservation when converting HTML to Excel with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Shows how to import an HTML file into an Aspose.Cells Workbook while preserving the original <div> layout by enabling SupportDivTag in HtmlLoadOptions, then export the result to XLSX.
    public class LoadHtmlWithDivTagLayout
    {
        public static void Main(string[] args)
        {
            Run();
        }

        public static void Run()
        {
            try
            {
                string inputPath = "input.html";
                string outputPath = "output.xlsx";

                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Create HTML load options and enable support for <div> tag layout
                HtmlLoadOptions loadOptions = new HtmlLoadOptions(LoadFormat.Html)
                {
                    SupportDivTag = true
                };

                // Load the HTML file into a workbook using the configured options
                Workbook workbook = new Workbook(inputPath, loadOptions);

                // Access the first worksheet to verify data was loaded
                Worksheet sheet = workbook.Worksheets[0];
                Console.WriteLine("First cell value: " + sheet.Cells["A1"].StringValue);

                // Save the workbook to an Excel file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
