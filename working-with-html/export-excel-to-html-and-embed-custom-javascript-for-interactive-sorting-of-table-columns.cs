// Title: Export Excel to HTML with Custom JavaScript Sorting – Aspose.Cells for .NET
// Description: Creates a workbook, fills it with sample data, configures HtmlSaveOptions to disable default frame scripts and output HTML5, saves the sheet to a memory stream, injects a lightweight JavaScript sorter before the closing </body> tag, and writes the final HTML file with clickable column headers that toggle ascending/descending order.
// Keywords: Aspose.Cells HTML export | custom JavaScript injection | interactive column sorting | disable frame scripts | HtmlSaveOptions .NET | HTML5 Excel output | client‑side table sort | C# Excel to HTML
// Common Searches: Aspose.Cells add JavaScript to exported HTML | how to enable column sorting in Aspose.Cells HTML output | remove default frame scripts when saving Excel as HTML | inject script before </body> in Aspose.Cells generated HTML | C# export Excel worksheet to sortable HTML page
// Developer Intent: Export an Excel worksheet to an HTML file and embed a custom JavaScript routine that lets end‑users sort table columns directly in the browser.
// Use Cases: Publish a product catalog from Excel with sortable price and quantity columns. | Create a lightweight web report without Aspose.Cells frame scripts, adding only the needed sorting logic. | Embed a single‑sheet HTML view in a web portal where users can reorder data client‑side.
// AI Prompts: Modify the injected script to correctly sort dates formatted as MM/dd/yyyy. | Show how to export multiple worksheets to separate HTML files while preserving custom JavaScript. | Provide a robust method to locate and replace the </body> tag when the generated HTML lacks a body element.

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsHtmlExportWithSorting
{
    // Creates a workbook, fills it with sample data, configures HtmlSaveOptions to disable default frame scripts and output HTML5, saves the sheet to a memory stream, injects a lightweight JavaScript sorter before the closing </body> tag, and writes the final HTML file with clickable column headers that toggle ascending/descending order.
    class Program
    {
        static void Main()
        {
            // -------------------------------------------------
            // 1. Create a workbook and populate it with sample data
            // -------------------------------------------------
            Workbook workbook = new Workbook();                     // create workbook
            Worksheet sheet = workbook.Worksheets[0];               // get first worksheet
            Cells cells = sheet.Cells;

            // Header row
            cells["A1"].PutValue("Product");
            cells["B1"].PutValue("Price");
            cells["C1"].PutValue("Quantity");

            // Sample data rows
            cells["A2"].PutValue("Laptop");
            cells["B2"].PutValue(1200);
            cells["C2"].PutValue(5);

            cells["A3"].PutValue("Phone");
            cells["B3"].PutValue(800);
            cells["C3"].PutValue(12);

            cells["A4"].PutValue("Tablet");
            cells["B4"].PutValue(450);
            cells["C4"].PutValue(8);

            // -------------------------------------------------
            // 2. Configure HTML save options
            // -------------------------------------------------
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            // Disable default frame scripts so we can inject our own JavaScript only
            htmlOptions.ExportFrameScriptsAndProperties = false;
            // Export only the active worksheet (optional)
            htmlOptions.ExportActiveWorksheetOnly = true;
            // Use HTML5 for modern browsers (optional)
            htmlOptions.HtmlVersion = HtmlVersion.Html5;

            // -------------------------------------------------
            // 3. Save the workbook to a memory stream as HTML
            // -------------------------------------------------
            using (MemoryStream htmlStream = new MemoryStream())
            {
                workbook.Save(htmlStream, htmlOptions);   // save using provided save rule
                htmlStream.Position = 0;                  // reset stream position

                // Read the generated HTML into a string
                string htmlContent = new StreamReader(htmlStream, Encoding.UTF8).ReadToEnd();

                // -------------------------------------------------
                // 4. Inject custom JavaScript for column sorting
                // -------------------------------------------------
                string sortingScript = @"
<script type='text/javascript'>
// Simple table column sorter
document.addEventListener('DOMContentLoaded', function () {
    var tables = document.getElementsByTagName('table');
    for (var t = 0; t < tables.length; t++) {
        makeSortable(tables[t]);
    }
});

function makeSortable(table) {
    var ths = table.getElementsByTagName('th');
    for (var i = 0; i < ths.length; i++) {
        (function (index) {
            ths[index].style.cursor = 'pointer';
            ths[index].addEventListener('click', function () {
                sortTable(table, index);
            });
        })(i);
    }
}

function sortTable(table, colIndex) {
    var rows = Array.prototype.slice.call(table.tBodies[0].rows, 0);
    var asc = table.getAttribute('data-sort-dir') !== 'asc';
    rows.sort(function (a, b) {
        var aText = a.cells[colIndex].textContent.trim();
        var bText = b.cells[colIndex].textContent.trim();
        var aNum = parseFloat(aText);
        var bNum = parseFloat(bText);
        if (!isNaN(aNum) && !isNaN(bNum)) {
            return asc ? aNum - bNum : bNum - aNum;
        }
        return asc ? aText.localeCompare(bText) : bText.localeCompare(aText);
    });
    // Re‑append sorted rows
    for (var i = 0; i < rows.length; i++) {
        table.tBodies[0].appendChild(rows[i]);
    }
    table.setAttribute('data-sort-dir', asc ? 'asc' : 'desc');
}
</script>";

                // Insert the script just before the closing </body> tag
                int bodyCloseIndex = htmlContent.LastIndexOf("</body>", StringComparison.OrdinalIgnoreCase);
                if (bodyCloseIndex >= 0)
                {
                    htmlContent = htmlContent.Insert(bodyCloseIndex, sortingScript);
                }
                else
                {
                    // Fallback: append at the end
                    htmlContent += sortingScript;
                }

                // -------------------------------------------------
                // 5. Write the final HTML with embedded script to a file
                // -------------------------------------------------
                File.WriteAllText("WorkbookWithSorting.html", htmlContent, Encoding.UTF8);
                Console.WriteLine("HTML file with interactive sorting saved as 'WorkbookWithSorting.html'.");
            }
        }
    }
}
