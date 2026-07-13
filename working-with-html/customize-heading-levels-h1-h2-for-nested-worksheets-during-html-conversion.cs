using System;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Saving;

class Program
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Configure HTML save options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            // Export row and column headings (equivalent to sheet headings)
            ExportRowColumnHeadings = true,
            // Export the whole workbook (default behavior)
            ExportActiveWorksheetOnly = false
        };

        // Save to a memory stream first to allow post‑processing
        using (MemoryStream ms = new MemoryStream())
        {
            workbook.Save(ms, htmlOptions);
            ms.Position = 0;
            string html = Encoding.UTF8.GetString(ms.ToArray());

            // Post‑process the HTML to insert <h1> tags for each worksheet.
            // Aspose.Cells generates a div with id equal to the worksheet name.
            // Replace that div start tag with an <h1> heading followed by the div.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                string divId = $"id=\"{sheet.Name}\"";
                string replacement = $"<h1>{sheet.Name}</h1><div {divId}";
                html = html.Replace($"<div {divId}", replacement);
            }

            // Write the modified HTML to the output file
            File.WriteAllText("output.html", html, Encoding.UTF8);
        }
    }
}