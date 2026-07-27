// Title: Export Excel to Responsive HTML with Aspose.Cells for .NET (C#)
// Description: Load an Excel workbook with Aspose.Cells, configure HtmlSaveOptions for HTML5, enable mobile compatibility, optionally export only the active sheet or embed all resources into a single file, and save a mobile‑friendly HTML document ready for smartphones and tablets.
// Keywords: Aspose.Cells | C# | Excel to HTML | responsive HTML export | HTML5 | IsMobileCompatible | single file HTML | mobile friendly | HtmlSaveOptions | export active worksheet
// Common Searches: Aspose.Cells export Excel to responsive HTML | C# generate mobile friendly HTML from Excel | HtmlSaveOptions IsMobileCompatible example | single file HTML export Aspose.Cells | how to add viewport meta tag with Aspose.Cells
// Developer Intent: Create a mobile‑optimized HTML page from an Excel workbook using Aspose.Cells in C#.
// Use Cases: Deliver a single‑file HTML report that scales correctly on smartphones. | Embed a lightweight worksheet view into a web portal without extra assets. | Generate HTML5 output with automatic viewport settings for cross‑device compatibility.
// AI Prompts: Write C# code that loads an Excel file and saves it as responsive HTML5 using Aspose.Cells, with IsMobileCompatible enabled and resources embedded in one file. | Explain how to configure HtmlSaveOptions to export only the active worksheet and produce a single HTML file for mobile consumption. | Provide troubleshooting steps when the Aspose.Cells generated HTML does not render correctly on iOS Safari.

using System;
using Aspose.Cells;

namespace AsposeCellsResponsiveHtmlExport
{
    // Load an Excel workbook with Aspose.Cells, configure HtmlSaveOptions for HTML5, enable mobile compatibility, optionally export only the active sheet or embed all resources into a single file, and save a mobile‑friendly HTML document ready for smartphones and tablets.
    class Program
    {
        static void Main()
        {
            // Load the source Excel workbook
            // (Replace the path with the actual location of your Excel file)
            string sourcePath = "input.xlsx";
            Workbook workbook = new Workbook(sourcePath);

            // Configure HTML save options for responsive (mobile‑friendly) output
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Use HTML5 which works well with modern browsers and mobile devices
            htmlOptions.HtmlVersion = HtmlVersion.Html5;

            // Enable built‑in mobile compatibility (adds viewport meta tag, etc.)
            htmlOptions.IsMobileCompatible = true;

            // Optional: export only the active worksheet to keep the HTML lightweight
            // htmlOptions.ExportActiveWorksheetOnly = true;

            // Optional: embed all resources into a single HTML file (CSS, images as Base64)
            // This helps when serving the file on mobile where multiple files are inconvenient
            // htmlOptions.SaveAsSingleFile = true;

            // Save the workbook as an HTML file with the above options
            // (Replace the path with the desired output location)
            string outputPath = "output_responsive.html";
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"Responsive HTML file saved to: {outputPath}");
        }
    }
}
