// Title: Merge cells across rows in an Excel worksheet and save as HTML using Aspose.Cells for .NET
// AI Prompts: Load an Excel file, merge the range A1:B3 on the first worksheet, and export the sheet to an HTML file with the default HtmlSaveOptions in C#. | Set HtmlSaveOptions to include all worksheet data and preserve merged cells, then save the workbook as HTML using Aspose.Cells.
// Common Searches: c# aspnet merge cells A1:B3 and export worksheet to html using Aspose.Cells | how to configure HtmlSaveOptions to export merged cells to HTML in Aspose.Cells .NET | Aspose.Cells example for merging a range of cells and saving as an HTML file | export Excel workbook with merged rows to HTML with all data using C# | Aspose.Cells merge cells across rows then convert workbook to HTML
// Tags: merge cells Aspose.Cells C# | HtmlSaveOptions ExportDataOptions All | Aspose.Cells HTML export of merged cells | C# load workbook save as HTML | merged range conversion to HTML Aspose.Cells

using System;
using Aspose.Cells;

namespace MergeAndExportToHtml
{
    // The program loads input.xlsx, merges the A1:B3 range on the first worksheet, configures HtmlSaveOptions to export all data while preserving merged cells, and saves the result as output.html.
    class Program
    {
        static void Main()
        {
            // Path to the source Excel file
            string inputFile = "input.xlsx";

            // Load the workbook from the file
            Workbook workbook = new Workbook(inputFile);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Merge cells across rows.
            // This merges the range A1:B3 (zero‑based indices: row 0, column 0, spanning 3 rows and 2 columns)
            worksheet.Cells.Merge(0, 0, 3, 2);

            // Configure HTML save options.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Merge contiguous empty TD elements (default behavior)
                MergeEmptyTdType = MergeEmptyTdType.Default,

                // Export the entire worksheet data (including headings)
                ExportDataOptions = HtmlExportDataOptions.All
            };

            // Path for the resulting HTML file
            string outputFile = "output.html";

            // Save the workbook as an HTML file using the specified options
            workbook.Save(outputFile, htmlOptions);

            Console.WriteLine($"Workbook merged and exported to HTML: {outputFile}");
        }
    }
}
