// Title: Export an Excel worksheet to HTML and add client‑side JavaScript for sortable columns with Aspose.Cells for .NET
// AI Prompts: Create C# code that loads a workbook with Aspose.Cells, saves the active sheet as HTML, and then injects a JavaScript block that enables click‑to‑sort on table header cells. | Write a routine that reads the generated HTML file, locates the closing </body> tag (or appends at the end if absent), and inserts a <script> element containing a simple table‑sorting function. | Add robust error handling to verify the source Excel file exists, ensure the output directory is created, and log any exceptions that occur during the HTML post‑processing step.
// Common Searches: how to make columns sortable in HTML exported from Excel using Aspose.Cells C# | Aspose.Cells export active worksheet to HTML and embed custom JavaScript | C# insert script tag into Aspose.Cells generated HTML before </body> | post‑process Aspose.Cells HTML output to add table sorting script | handle missing </body> tag when appending JavaScript to generated HTML in .NET
// Tags: Aspose.Cells export worksheet to HTML | inject JavaScript into Aspose.Cells HTML output | client‑side table sorting script C# | post‑process generated HTML file .NET | handle missing body tag when appending script

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

// The example loads an Excel file with Aspose.Cells, saves the active worksheet as an HTML page, then reads the HTML, inserts a JavaScript block that attaches click handlers to table headers for ascending/descending sorting, and writes the modified content back. It creates the output directory if needed, checks for the source file, and gracefully handles missing </body> tags or other I/O errors.
class ExcelToHtmlWithSorting
{
    static void Main()
    {
        try
        {
            // Path to the source Excel file
            string excelPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(excelPath))
            {
                Console.WriteLine($"Error: The file \"{excelPath}\" was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(excelPath);

            // Configure HTML save options
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                ExportActiveWorksheetOnly = true // export only the active sheet
                // IsFullHtml and CustomScripts are not available in the current API version,
                // so we will embed the custom script manually after saving.
            };

            // Path for the generated HTML file
            string htmlPath = "output.html";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(htmlPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as HTML
            workbook.Save(htmlPath, saveOptions);

            // Embed custom JavaScript for column sorting
            const string script = @"
<script type='text/javascript'>
// Simple table sorting script
document.addEventListener('DOMContentLoaded', function () {
    var table = document.querySelector('table');
    if (!table) return;

    var headers = table.querySelectorAll('th');
    for (let i = 0; i < headers.length; i++) {
        (function (index) {
            headers[index].addEventListener('click', function () {
                sortTable(table, index);
            });
        })(i);
    }

    function sortTable(tbl, colIndex) {
        var rows = Array.from(tbl.rows).slice(1); // exclude header row
        var asc = tbl.getAttribute('data-sort-dir') !== 'asc';
        rows.sort(function (a, b) {
            var aText = a.cells[colIndex].textContent.trim();
            var bText = b.cells[colIndex].textContent.trim();

            // Attempt numeric comparison
            var aNum = parseFloat(aText);
            var bNum = parseFloat(bText);
            if (!isNaN(aNum) && !isNaN(bNum)) {
                return asc ? aNum - bNum : bNum - aNum;
            }

            // Fallback to string comparison
            return asc ? aText.localeCompare(bText) : bText.localeCompare(aText);
        });

        // Re-append sorted rows
        for (let i = 0; i < rows.length; i++) {
            tbl.tBodies[0].appendChild(rows[i]);
        }

        // Store sort direction for next click
        tbl.setAttribute('data-sort-dir', asc ? 'asc' : 'desc');
    }
});
</script>";

            // Insert the script before the closing </body> tag
            try
            {
                string htmlContent = File.ReadAllText(htmlPath);
                int bodyCloseIndex = htmlContent.LastIndexOf("</body>", StringComparison.OrdinalIgnoreCase);
                if (bodyCloseIndex >= 0)
                {
                    htmlContent = htmlContent.Insert(bodyCloseIndex, script);
                    File.WriteAllText(htmlPath, htmlContent);
                }
                else
                {
                    // If </body> not found, append at the end
                    File.AppendAllText(htmlPath, script);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Unable to embed custom script. {ex.Message}");
            }

            Console.WriteLine("Excel file has been exported to HTML with interactive sorting.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
