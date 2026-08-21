// Title: Export XLS to HTML with All Cell Comments Using Aspose.Cells for .NET
// Description: Shows how to load an XLS workbook with Aspose.Cells, set HtmlSaveOptions.IsExportComments to true, and save the file as HTML so that every cell comment appears in the generated web page.
// Keywords: Aspose.Cells HTML export comments | C# convert XLS to HTML | IsExportComments option | preserve Excel comments .NET | export Excel to web page
// Common Searches: Aspose.Cells export cell comments to HTML C# | HtmlSaveOptions IsExportComments example | convert legacy XLS to HTML with comments | save Excel workbook as HTML preserving comments
// Developer Intent: Create an HTML file from an XLS workbook that includes all embedded cell comments.
// Use Cases: Publish legacy Excel reports on a website while keeping reviewer notes visible. | Generate documentation that shows data together with its comment annotations for end‑users. | Automate batch conversion of multiple XLS files to HTML, ensuring comment fidelity.
// AI Prompts: Write C# code with Aspose.Cells to load an XLS file and export it to HTML with every cell comment included. | Explain how to configure HtmlSaveOptions to export comments and optionally retain column widths and styles during HTML conversion. | Provide a script that processes a folder of XLS workbooks, converting each to HTML while preserving all comments using Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsExportComments
{
    // Shows how to load an XLS workbook with Aspose.Cells, set HtmlSaveOptions.IsExportComments to true, and save the file as HTML so that every cell comment appears in the generated web page.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source Excel file (XLS format)
            string inputPath = "input.xls";

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Create HTML save options and enable exporting of comments
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                IsExportComments = true // Ensure all cell comments are included in the HTML output
            };

            // Path for the generated HTML file
            string outputPath = "output.html";

            // Save the workbook as HTML using the configured options
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"HTML file with comments exported successfully to: {outputPath}");
        }
    }
}
