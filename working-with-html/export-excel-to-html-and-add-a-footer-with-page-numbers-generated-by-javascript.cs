// Title: How to export an Excel workbook to HTML with a fixed footer showing page numbers using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, saves it as HTML, and appends a fixed‑position footer element containing a page‑number placeholder. | Generate a JavaScript snippet that computes total pages from the document height and updates a span inside the footer with the current page number. | Show how to insert the footer HTML and the JavaScript block just before the closing </body> tag of the HTML string returned by Aspose.Cells.
// Common Searches: Aspose.Cells .NET export Excel to HTML and add custom footer with page numbers | C# insert JavaScript pagination into HTML generated from Excel using Aspose.Cells | How to add a fixed footer displaying page count when converting XLSX to HTML in .NET | Calculate total pages from document height in JavaScript for Aspose.Cells HTML output
// Tags: Aspose.Cells HTML export custom footer | C# inject JavaScript pagination | fixed footer page numbers | insert script before body close tag | document height page calculation JavaScript

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// Loads an Excel workbook, saves it as HTML using Aspose.Cells, injects a fixed‑position footer and a JavaScript block that approximates page numbers based on page height, then writes the modified HTML to an output file.
class ExcelToHtmlWithFooter
{
    static void Main()
    {
        // Load the source Excel workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Configure HTML save options (default options are sufficient for this task)
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

        // Save the workbook to a memory stream as HTML
        using (MemoryStream htmlStream = new MemoryStream())
        {
            workbook.Save(htmlStream, htmlOptions);
            htmlStream.Position = 0;

            // Convert the HTML bytes to a string (UTF-8 encoding)
            string htmlContent = Encoding.UTF8.GetString(htmlStream.ToArray());

            // Define the footer HTML that will hold the page number
            string footerHtml = @"
<div id='footer' style='position:fixed; bottom:0; left:0; width:100%; text-align:center; background:#f0f0f0; padding:5px; font-family:Arial,Helvetica,sans-serif;'>
    Page <span id='pageNum'></span>
</div>";

            // JavaScript that calculates a simple page number based on document height.
            // This is a basic example; for more accurate pagination you may need a more sophisticated approach.
            string scriptHtml = @"
<script type='text/javascript'>
document.addEventListener('DOMContentLoaded', function () {
    // Approximate page height in pixels (A4 @ 96 DPI ≈ 1122px)
    var pageHeight = 1122;
    var totalHeight = document.body.scrollHeight;
    var totalPages = Math.ceil(totalHeight / pageHeight);
    // For demonstration, show the first page out of total pages
    document.getElementById('pageNum').innerText = '1 of ' + totalPages;
});
</script>";

            // Insert the footer and script just before the closing </body> tag
            int bodyCloseIndex = htmlContent.LastIndexOf("</body>", StringComparison.OrdinalIgnoreCase);
            if (bodyCloseIndex >= 0)
            {
                string insertion = footerHtml + scriptHtml + Environment.NewLine;
                htmlContent = htmlContent.Insert(bodyCloseIndex, insertion);
            }
            else
            {
                // Fallback: append at the end if </body> not found
                htmlContent += footerHtml + scriptHtml;
            }

            // Write the modified HTML to the output file
            File.WriteAllText("output.html", htmlContent, Encoding.UTF8);
        }

        Console.WriteLine("Excel exported to HTML with footer page numbers.");
    }
}
