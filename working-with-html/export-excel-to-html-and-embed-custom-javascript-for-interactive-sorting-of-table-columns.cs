// Title: Convert an Excel workbook to HTML with Aspose.Cells and add a client‑side column‑sorting script in C#
// AI Prompts: Write C# code that loads an .xlsx file using Aspose.Cells, saves it as HTML with HtmlSaveOptions (embedding images as Base64), and writes the file to disk. | Implement a helper method that reads the generated HTML, finds the closing </body> tag case‑insensitively, and injects a JavaScript block that enables click‑to‑sort on all table columns, appending the script if the tag is absent. | Provide a lightweight JavaScript snippet that attaches click handlers to <th> elements to sort their column in ascending or descending order.
// Common Searches: how to export Excel to HTML with Aspose.Cells and include a custom JavaScript sorter | c# inject JavaScript into Aspose.Cells generated HTML before the closing body tag | Aspose.Cells HtmlSaveOptions embed images as base64 and add client side table sorting | add click to sort functionality to tables in HTML produced from an Excel workbook using C# | save workbook as HTML then programmatically insert a script for column sorting
// Tags: Aspose.Cells HTML export with custom script | C# inject JavaScript into generated HTML | Excel to HTML conversion base64 images | client‑side table column sorter JavaScript | post‑process Aspose.Cells HTML output

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel file with Aspose.Cells, saves it as an HTML document using HtmlSaveOptions that embed images as Base64, and then reads the HTML file to inject a JavaScript snippet before the closing </body> tag (or appends it if the tag is missing). The injected script adds click‑to‑sort behavior to every table column, providing interactive sorting in the generated HTML.
class ExcelToHtmlWithSorting
{
    static void Main()
    {
        try
        {
            string inputPath = "Input.xlsx";
            string outputPath = "Output.html";

            // Ensure the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options (without CustomScripts)
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                ExportActiveWorksheetOnly = false,   // Export entire workbook
                ExportImagesAsBase64 = true          // Embed images as Base64
            };

            // Save the workbook as HTML
            workbook.Save(outputPath, saveOptions);
            Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");

            // Inject sorting script into the generated HTML
            InjectSortingScript(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Inserts the JavaScript snippet before the closing </body> tag
    private static void InjectSortingScript(string htmlPath)
    {
        try
        {
            if (!File.Exists(htmlPath))
            {
                Console.WriteLine($"HTML file '{htmlPath}' not found for script injection.");
                return;
            }

            string htmlContent = File.ReadAllText(htmlPath);
            string script = GetSortingScript();

            // Find the closing </body> tag (case‑insensitive) and insert the script before it
            int bodyCloseIndex = htmlContent.LastIndexOf("</body>", StringComparison.OrdinalIgnoreCase);
            if (bodyCloseIndex >= 0)
            {
                htmlContent = htmlContent.Insert(bodyCloseIndex, script + Environment.NewLine);
                File.WriteAllText(htmlPath, htmlContent);
                Console.WriteLine("Sorting script injected into HTML.");
            }
            else
            {
                // If </body> not found, append the script at the end
                File.AppendAllText(htmlPath, Environment.NewLine + script);
                Console.WriteLine("Sorting script appended to HTML (no </body> tag found).");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to inject sorting script: {ex.Message}");
        }
    }

    // Returns a JavaScript snippet that adds click‑to‑sort functionality to all tables
    private static string GetSortingScript()
    {
        return @"
<script type='text/javascript'>
// Simple table column sorter
document.addEventListener('DOMContentLoaded', function () {
    var tables = document.querySelectorAll('table');
    tables.forEach(function (table) {
        var headers = table.querySelectorAll('th');
        headers.forEach(function (header, index) {
            header.style.cursor = 'pointer';
            header.addEventListener('click', function () {
                sortTableByColumn(table, index);
            });
        });
    });
});

function sortTableByColumn(table, columnIndex) {
    var tbody = table.tBodies[0];
    var rows = Array.from(tbody.querySelectorAll('tr'));

    var asc = table.getAttribute('data-sort-dir') !== 'asc';
    table.setAttribute('data-sort-dir', asc ? 'asc' : 'desc');

    rows.sort(function (a, b) {
        var aText = a.cells[columnIndex].textContent.trim();
        var bText = b.cells[columnIndex].textContent.trim();

        var aNum = parseFloat(aText.replace(/[^0-9.-]+/g, ''));
        var bNum = parseFloat(bText.replace(/[^0-9.-]+/g, ''));

        if (!isNaN(aNum) && !isNaN(bNum)) {
            return asc ? aNum - bNum : bNum - aNum;
        }

        return asc ? aText.localeCompare(bText) : bText.localeCompare(aText);
    });

    rows.forEach(function (row) {
        tbody.appendChild(row);
    });
}
</script>";
    }
}
