// Title: Aspose.Cells C# – Export Large Numbers to HTML as Plain Text (No Scientific Notation)
// Description: Creates a workbook, writes a large double (1234567890123456) to cell A1, applies the custom format "0" to suppress scientific notation, saves the sheet as HTML, reads the file and confirms that the output contains the number in plain decimal form, then prints the <td> element for manual review.
// Keywords: Aspose.Cells | HTML export | large number formatting | prevent scientific notation | custom number format 0 | C# | verify HTML output | exponential notation detection | cell style | HtmlSaveOptions
// Common Searches: Aspose.Cells stop scientific notation in HTML | C# export large numbers to HTML without exponent | custom number format for large values Aspose.Cells | check HTML output for exponent notation Aspose | verify plain number display in Aspose.Cells HTML
// Developer Intent: Confirm that a large numeric value is rendered in the generated HTML as a regular decimal string, not in exponential form.
// Use Cases: Apply the "0" custom format to a cell before HTML conversion to force plain representation of big numbers. | Read the saved HTML file and search for "E" or "e" to detect unwanted scientific notation. | Extract and display the <td> element containing the value for quick visual verification.
// AI Prompts: Generate C# code using Aspose.Cells that exports a worksheet to HTML while keeping large numbers in plain decimal format. | Provide a method to validate an Aspose.Cells‑generated HTML file, ensuring it does not contain exponent symbols for a specific cell. | Explain how to configure HtmlSaveOptions and cell styling to avoid exponential notation when saving large numbers to HTML with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlNumberVerification
{
    // Creates a workbook, writes a large double (1234567890123456) to cell A1, applies the custom format "0" to suppress scientific notation, saves the sheet as HTML, reads the file and confirms that the output contains the number in plain decimal form, then prints the <td> element for manual review.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Large numeric value that could be displayed in exponential form
                double largeNumber = 1234567890123456.0;

                // Put the numeric value into cell A1
                Cell cell = sheet.Cells["A1"];
                cell.PutValue(largeNumber);

                // Apply a custom number format to force plain decimal representation
                // "0" format displays the number without scientific notation
                Style style = cell.GetStyle();
                style.Custom = "0";
                cell.SetStyle(style);

                // Configure HTML save options (default options are sufficient for this case)
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

                // Save the workbook as HTML
                string htmlPath = "LargeNumberOutput.html";
                workbook.Save(htmlPath, htmlOptions);

                // Ensure the HTML file was created before attempting to read it
                if (!File.Exists(htmlPath))
                {
                    Console.WriteLine($"Error: HTML file '{htmlPath}' was not found.");
                    return;
                }

                // Read the generated HTML file
                string htmlContent = File.ReadAllText(htmlPath);

                // Verify that the large number appears as plain text (no 'E' or 'e' for exponent)
                bool containsExponent = htmlContent.Contains("E") || htmlContent.Contains("e");
                Console.WriteLine("HTML contains exponent notation: " + containsExponent);

                // Output the relevant HTML snippet for manual inspection
                Console.WriteLine("HTML snippet containing the number:");
                int start = htmlContent.IndexOf("<td", StringComparison.OrdinalIgnoreCase);
                if (start >= 0)
                {
                    int end = htmlContent.IndexOf("</td>", start, StringComparison.OrdinalIgnoreCase);
                    if (end >= 0)
                    {
                        end += "</td>".Length;
                        Console.WriteLine(htmlContent.Substring(start, end - start));
                    }
                    else
                    {
                        Console.WriteLine("Closing </td> tag not found.");
                    }
                }
                else
                {
                    Console.WriteLine("<td> tag not found in the HTML output.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An unexpected error occurred: " + ex.Message);
            }
        }
    }
}
