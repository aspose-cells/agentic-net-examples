// Title: C# – Convert HTML with Inline CSS to Excel and Auto‑Adjust Row Height by Font Size using Aspose.Cells
// Description: Loads an HTML file while preserving inline styles, creates an Aspose.Cells workbook, scans each used row to find the largest font size, applies a 1.2 conversion factor to set the row height, and saves the result as an XLSX document.
// Keywords: Aspose.Cells HTML to Excel conversion | C# inline CSS to Excel | map CSS font-size to Excel row height | auto adjust row height Aspose.Cells | .NET export HTML table to XLSX | font size based row height scaling
// Common Searches: how to keep inline CSS when converting HTML to Excel with Aspose.Cells | set Excel row height from maximum font size in a row C# | convert HTML file to .xlsx and auto‑scale row heights | Aspose.Cells example mapping font-size to row height
// Developer Intent: Transform an HTML document with inline CSS into an Excel workbook and automatically set each row’s height according to the biggest font size in that row.
// Use Cases: Create printable Excel reports from styled HTML invoices where row dimensions match the text size. | Export web‑based tables to Excel while preserving visual layout by scaling rows to the largest font in each line. | Automate conversion of HTML email templates to Excel, keeping inline formatting and ensuring readable row spacing.
// AI Prompts: Generate a C# Aspose.Cells snippet that loads an HTML file with inline styles and sets row heights based on the maximum font size per row using a 1.2 conversion factor. | Explain how to modify the font‑size‑to‑row‑height factor for different display requirements when converting HTML to Excel with Aspose.Cells. | Provide best‑practice error‑handling patterns for adjusting row heights during HTML‑to‑Excel conversion in C#.

using System;
using System.IO;
using Aspose.Cells;

// Loads an HTML file while preserving inline styles, creates an Aspose.Cells workbook, scans each used row to find the largest font size, applies a 1.2 conversion factor to set the row height, and saves the result as an XLSX document.
class HtmlToExcelConverter
{
    static void Main()
    {
        const string inputPath = "input.html";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input HTML file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the HTML file (inline styles are preserved)
            var loadOptions = new HtmlLoadOptions();
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Work with the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Determine the used range
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            // Adjust each row height based on the largest font size in that row
            for (int row = 0; row <= maxRow; row++)
            {
                double maxFontSize = 0.0;

                // Scan all columns in the current row
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = cells[row, col];
                    if (cell != null && cell.Value != null)
                    {
                        double fontSize = cell.GetStyle().Font.Size;
                        if (fontSize > maxFontSize)
                            maxFontSize = fontSize;
                    }
                }

                // If any font size was found, set the row height proportionally
                if (maxFontSize > 0)
                {
                    try
                    {
                        // Simple conversion: height = font size * 1.2 (points to pixels approximation)
                        sheet.Cells.Rows[row].Height = maxFontSize * 1.2;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Warning: Unable to set height for row {row}. {ex.Message}");
                    }
                }
            }

            // Save the workbook as an Excel file
            workbook.Save(outputPath);
            Console.WriteLine($"Conversion completed. Output saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
