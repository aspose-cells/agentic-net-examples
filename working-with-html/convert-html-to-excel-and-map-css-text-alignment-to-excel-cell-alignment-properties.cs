// Title: C# – Convert HTML Table to Excel with Aspose.Cells while preserving CSS text‑align
// Description: Loads an HTML file using Aspose.Cells, reads the inline style of each <td>, extracts the text‑align property, translates it to Aspose.Cells TextAlignmentType, applies the horizontal alignment to the matching worksheet cell, and saves the result as an XLSX workbook.
// Keywords: Aspose.Cells HTML to Excel | C# HTML table conversion | preserve CSS alignment in Excel | text-align to cell style mapping | loadOptions LoadFormat.Html | extract inline style attribute | horizontal alignment Aspose.Cells | convert HTML report to XLSX
// Common Searches: how to keep CSS text-align when converting HTML to Excel with Aspose.Cells | C# read td style attribute and set cell alignment | Aspose.Cells map inline CSS to Excel cell formatting | convert HTML table to XLSX preserving column alignment
// Developer Intent: Translate CSS text‑align values from HTML table cells into the corresponding horizontal alignment of Excel cells using Aspose.Cells.
// Use Cases: Migrate a web‑based report that uses left, center, and right aligned columns into a spreadsheet that looks identical. | Generate an analysis‑ready Excel file from an HTML email template while keeping the original cell alignment. | Automate bulk conversion of legacy HTML tables to XLSX files without losing visual layout defined by inline CSS.
// AI Prompts: Write C# code that loads an HTML file with Aspose.Cells, extracts the text‑align style from each <td>, maps it to TextAlignmentType, and applies it to the worksheet cells. | Enhance the converter to safely handle missing style attributes and add support for CSS values such as 'start', 'end', and 'inherit'. | Explain how extending MapCssAlignmentToTextAlignment to include 'justify' and 'initial' affects the resulting Excel cell formatting.

using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Aspose.Cells;

// Loads an HTML file using Aspose.Cells, reads the inline style of each <td>, extracts the text‑align property, translates it to Aspose.Cells TextAlignmentType, applies the horizontal alignment to the matching worksheet cell, and saves the result as an XLSX workbook.
class HtmlToExcelConverter
{
    static void Main()
    {
        try
        {
            // Path to the source HTML file
            string htmlPath = "input.html";

            // Verify that the HTML file exists to avoid FileNotFoundException
            if (!File.Exists(htmlPath))
            {
                Console.WriteLine($"Error: The file '{htmlPath}' was not found.");
                return;
            }

            // Load the HTML file into a workbook using Aspose.Cells LoadOptions for HTML format
            var loadOptions = new Aspose.Cells.LoadOptions(LoadFormat.Html);
            Workbook workbook = new Workbook(htmlPath, loadOptions);

            // Parse the HTML to extract CSS text‑align values from <td> elements
            XDocument doc = XDocument.Load(htmlPath);
            var rows = doc.Descendants("tr").ToList();

            // Iterate through rows and cells in the same order as they appear in the worksheet
            for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++)
            {
                var cellsInRow = rows[rowIndex].Descendants("td").ToList();
                for (int colIndex = 0; colIndex < cellsInRow.Count; colIndex++)
                {
                    // Get the style attribute of the current <td>
                    string styleAttr = cellsInRow[colIndex].Attribute("style")?.Value ?? string.Empty;

                    // Extract the value of text-align if present
                    string alignValue = ExtractTextAlign(styleAttr);

                    // Map CSS alignment to Aspose.Cells TextAlignmentType
                    TextAlignmentType alignment = MapCssAlignmentToTextAlignment(alignValue);

                    // Apply the alignment to the corresponding cell in the worksheet
                    Cell cell = workbook.Worksheets[0].Cells[rowIndex, colIndex];
                    Style cellStyle = cell.GetStyle();
                    cellStyle.HorizontalAlignment = alignment;
                    cell.SetStyle(cellStyle);
                }
            }

            // Save the workbook as an Excel file
            string excelPath = "output.xlsx";
            workbook.Save(excelPath);
            Console.WriteLine($"Conversion completed successfully. Excel file saved as '{excelPath}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Helper method to extract the value of text-align from a style string
    private static string ExtractTextAlign(string style)
    {
        // Example style: "color:#000000; text-align:center; font-weight:bold;"
        if (string.IsNullOrEmpty(style))
            return string.Empty;

        var parts = style.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
        foreach (var part in parts)
        {
            var kv = part.Split(new[] { ':' }, 2);
            if (kv.Length == 2 && kv[0].Trim().Equals("text-align", StringComparison.OrdinalIgnoreCase))
                return kv[1].Trim().ToLowerInvariant();
        }
        return string.Empty;
    }

    // Helper method to map CSS alignment strings to Aspose.Cells TextAlignmentType
    private static TextAlignmentType MapCssAlignmentToTextAlignment(string cssAlign)
    {
        return cssAlign switch
        {
            "left" => TextAlignmentType.Left,
            "center" => TextAlignmentType.Center,
            "right" => TextAlignmentType.Right,
            "justify" => TextAlignmentType.Justify,
            _ => TextAlignmentType.General,
        };
    }
}
