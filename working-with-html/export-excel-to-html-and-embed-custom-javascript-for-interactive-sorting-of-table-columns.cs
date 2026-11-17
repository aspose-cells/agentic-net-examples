// Title: Convert Excel workbook to HTML with Aspose.Cells and embed client‑side JavaScript for sortable columns (C#)
// Common Searches: how to export an .xlsx file to HTML with Aspose.Cells and include custom JavaScript for column sorting | C# Aspose.Cells HTML export inject JavaScript before </head> tag | add client side table sorting to HTML generated from Excel using Aspose.Cells .NET | Aspose.Cells save workbook as HTML and embed sortable table script | inject JavaScript into Aspose.Cells HTML output for interactive column headers
// Tags: Aspose.Cells HtmlSaveOptions with custom script | C# generate sortable HTML table from Excel workbook | inject JavaScript into Aspose.Cells HTML output | client‑side column sorting for exported Excel HTML | HTML export of .xlsx with interactive table headers

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// The example loads an existing .xlsx file with Aspose.Cells, saves it to HTML using HtmlSaveOptions into a memory stream, converts the stream to a UTF‑8 string, defines a JavaScript function that sorts table rows when header cells are clicked, inserts the script before the closing </head> tag (or appends it if the tag is missing), and writes the final HTML file. It includes file‑existence checks and generic exception handling.
class ExcelToHtmlWithSorting
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.html";

        try
        {
            // Verify that the source Excel file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the source Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

            // Save the workbook to a memory stream first
            using (MemoryStream ms = new MemoryStream())
            {
                workbook.Save(ms, htmlOptions);
                ms.Position = 0;

                // Convert the generated HTML to a string
                string htmlContent = Encoding.UTF8.GetString(ms.ToArray());

                // JavaScript for interactive column sorting
                string sortingScript = @"
<script type='text/javascript'>
function sortTable(colIndex) {
    var table = document.getElementsByTagName('table')[0];
    var rows, switching, i, x, y, shouldSwitch, dir = 'asc', switchcount = 0;
    switching = true;
    while (switching) {
        switching = false;
        rows = table.rows;
        for (i = 1; i < (rows.length - 1); i++) {
            shouldSwitch = false;
            x = rows[i].getElementsByTagName('TD')[colIndex];
            y = rows[i + 1].getElementsByTagName('TD')[colIndex];
            if (dir == 'asc') {
                if (x.innerHTML.toLowerCase() > y.innerHTML.toLowerCase()) {
                    shouldSwitch = true;
                    break;
                }
            } else if (dir == 'desc') {
                if (x.innerHTML.toLowerCase() < y.innerHTML.toLowerCase()) {
                    shouldSwitch = true;
                    break;
                }
            }
        }
        if (shouldSwitch) {
            rows[i].parentNode.insertBefore(rows[i + 1], rows[i]);
            switching = true;
            switchcount++;
        } else {
            if (switchcount == 0 && dir == 'asc') {
                dir = 'desc';
                switching = true;
            }
        }
    }
}

// Attach click events to header cells after the document loads
document.addEventListener('DOMContentLoaded', function () {
    var headers = document.getElementsByTagName('th');
    for (let i = 0; i < headers.length; i++) {
        headers[i].style.cursor = 'pointer';
        headers[i].addEventListener('click', (function (index) {
            return function () { sortTable(index); };
        })(i));
    }
});
</script>";

                // Insert the script just before the closing </head> tag; if not found, append at the end
                int headCloseIndex = htmlContent.IndexOf("</head>", StringComparison.OrdinalIgnoreCase);
                if (headCloseIndex >= 0)
                {
                    htmlContent = htmlContent.Insert(headCloseIndex, sortingScript);
                }
                else
                {
                    htmlContent += sortingScript;
                }

                // Write the final HTML to the output file
                File.WriteAllText(outputPath, htmlContent, Encoding.UTF8);
                Console.WriteLine($"HTML file with sorting script saved to '{outputPath}'.");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions to prevent the program from crashing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
