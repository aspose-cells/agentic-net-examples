// Title: C# – Export Excel to HTML with UTF‑8 Encoding using Aspose.Cells
// Description: Loads an Excel workbook, configures HtmlSaveOptions with Encoding = Encoding.UTF8, and saves the file as HTML. The UTF‑8 setting guarantees correct rendering of multilingual and special characters in the generated web page.
// Keywords: Aspose.Cells | C# | .NET | Excel to HTML | UTF-8 encoding | HtmlSaveOptions | Unicode support | multilingual export | web report generation | Workbook.Save HTML
// Common Searches: Aspose.Cells export Excel to HTML UTF-8 | C# set HtmlSaveOptions encoding to UTF-8 | How to preserve Unicode characters when converting Excel to HTML | Save workbook as HTML with international character support | Aspose.Cells HTML conversion encoding options
// Developer Intent: Create an HTML version of an Excel workbook that uses UTF‑8 encoding to display international characters correctly.
// Use Cases: Publish multilingual spreadsheets as web‑ready reports. | Email HTML snapshots of Excel data without character corruption. | Integrate Excel‑derived content into web applications that require Unicode compliance.
// AI Prompts: Generate C# code that converts an Excel file to a single‑file HTML document with UTF‑8 encoding and embedded images using Aspose.Cells. | Explain how to validate the character encoding of the saved HTML file after using HtmlSaveOptions. | Show how to combine UTF‑8 encoding with additional HtmlSaveOptions such as custom CSS, inline images, and page layout settings.

using System;
using System.Text;
using Aspose.Cells;

// Loads an Excel workbook, configures HtmlSaveOptions with Encoding = Encoding.UTF8, and saves the file as HTML. The UTF‑8 setting guarantees correct rendering of multilingual and special characters in the generated web page.
class ExportExcelToHtmlUtf8
{
    static void Main()
    {
        // Load an existing Excel workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Create HTML save options
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();

        // Set the output encoding to UTF‑8 for proper international character support
        saveOptions.Encoding = Encoding.UTF8;

        // Save the workbook as an HTML file using the specified options
        workbook.Save("output.html", saveOptions);
    }
}
