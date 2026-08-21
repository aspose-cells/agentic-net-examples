// Title: Aspose.Cells HtmlSaveOptions: Export Numeric Cells as Plain Text (No Scientific Notation) in C#/.NET
// Description: This C# example shows how to stop scientific notation when converting an Excel workbook to HTML with Aspose.Cells. It creates a workbook, adds a large integer and a tiny decimal, applies a simple number style (Number = 0) to the cells, sets HtmlSaveOptions to export the displayed values and ignore column width, and saves the output so the numbers are rendered as ordinary text.
// Keywords: Aspose.Cells | HtmlSaveOptions | C# | .NET | export numeric cells to HTML | disable scientific notation | plain number format | displayed values only | ignore column width | Excel to HTML conversion | cell style Number property
// Common Searches: Aspose.Cells prevent scientific notation in HTML | HtmlSaveOptions export displayed values C# | format cells as plain numbers before HTML export | C# Aspose.Cells export numbers without exponent | show full numeric value in HTML using Aspose.Cells
// Developer Intent: Create HTML output where numeric cells are rendered as ordinary text instead of scientific notation.
// Use Cases: Web dashboards that need exact numeric representation without exponents | Financial statements or invoices displayed in browsers with human‑readable numbers | Regulatory data exports where scientific notation is prohibited | Embedding Excel data in web pages while preserving original formatting | Automated HTML reporting pipelines that require clear numeric values
// AI Prompts: Provide C# code using Aspose.Cells to export a workbook to HTML with numbers shown as ordinary text, not scientific notation. | Show how to apply a non‑exponential number style to a range and configure HtmlSaveOptions to export displayed values only. | Explain which HtmlSaveOptions properties keep column width ignored and preserve cell formatting during Excel‑to‑HTML conversion.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsHtmlExport
{
    // This C# example shows how to stop scientific notation when converting an Excel workbook to HTML with Aspose.Cells. It creates a workbook, adds a large integer and a tiny decimal, applies a simple number style (Number = 0) to the cells, sets HtmlSaveOptions to export the displayed values and ignore column width, and saves the output so the numbers are rendered as ordinary text.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate cells with numeric values that would normally appear in scientific notation
                sheet.Cells["A1"].PutValue(123456789);      // Large integer
                sheet.Cells["A2"].PutValue(0.00000123);    // Small decimal

                // Apply a plain number format to the cells to force non‑scientific display
                Style plainStyle = workbook.CreateStyle();
                plainStyle.Number = 0; // Integer format without exponent

                // Apply the style to the range A1:A2
                AsposeRange range = sheet.Cells.CreateRange("A1:A2");
                range.ApplyStyle(plainStyle, new StyleFlag { All = true });

                // Configure HtmlSaveOptions
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    // Export displayed values (not raw values)
                    ExportDataOptions = HtmlExportDataOptions.All,
                    // Ignore column width so the full formatted text is shown
                    FormatDataIgnoreColumnWidth = true
                };

                // Save the workbook as HTML using the configured options
                workbook.Save("NumericPlainText.html", htmlOptions);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
