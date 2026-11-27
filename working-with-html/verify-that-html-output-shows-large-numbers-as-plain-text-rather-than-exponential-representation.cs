// Title: C# example: Verify large integer is rendered as plain text (no scientific notation) in HTML saved with Aspose.Cells
// Common Searches: C# Aspose.Cells export large integer to HTML without scientific notation | how to check that a 19‑digit number appears as plain text in Aspose.Cells HTML output | prevent exponential format for big numbers when saving workbook as HTML using Aspose.Cells .NET | read generated HTML file and confirm exact number string with Aspose.Cells C# | Aspose.Cells HtmlSaveOptions keep large numbers as text in C#
// Tags: Aspose.Cells large number HTML export | text format cell style Aspose.Cells | C# read generated HTML verification | HtmlSaveOptions usage Aspose.Cells | prevent exponential notation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace LargeNumberHtmlVerification
{
    // The sample creates a workbook, writes a 19‑digit number as a string into cell A1, applies a text style (Number = 0, Custom = "@"), saves the workbook to HTML with default HtmlSaveOptions, reads the resulting HTML file, and confirms that the exact number appears in the markup while no "E+" scientific notation is present, reporting success or failure.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Prepare output directory
                string dataDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
                Directory.CreateDirectory(dataDir);
                string htmlPath = Path.Combine(dataDir, "LargeNumberOutput.html");

                // Large number as string to keep precision
                const string largeNumberString = "1234567890123456789";

                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Put the large number into cell A1 as text
                Cell cell = worksheet.Cells["A1"];
                cell.PutValue(largeNumberString);

                // Apply a text format to the cell
                Style textStyle = workbook.CreateStyle();
                textStyle.Number = 0;          // General number format
                textStyle.Custom = "@";        // Text format
                cell.SetStyle(textStyle);

                // Configure HTML save options (no special property needed for plain text)
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

                // Save the workbook as HTML
                workbook.Save(htmlPath, htmlOptions);
                Console.WriteLine($"Workbook saved as HTML to: {htmlPath}");

                // Verify the HTML file was created
                if (!File.Exists(htmlPath))
                {
                    Console.WriteLine("Error: HTML file was not created.");
                    return;
                }

                // Read the generated HTML content
                string htmlContent;
                try
                {
                    htmlContent = File.ReadAllText(htmlPath);
                }
                catch (Exception readEx)
                {
                    Console.WriteLine($"Failed to read HTML file: {readEx.Message}");
                    return;
                }

                // Check for plain number and absence of exponential notation
                bool containsPlainNumber = htmlContent.Contains(largeNumberString);
                bool containsExponential = htmlContent.IndexOf("E+", StringComparison.OrdinalIgnoreCase) >= 0;

                Console.WriteLine();
                Console.WriteLine("Verification Results:");
                if (containsPlainNumber && !containsExponential)
                {
                    Console.WriteLine("SUCCESS: The large number is present in the HTML as plain text.");
                }
                else
                {
                    Console.WriteLine("FAILURE: The large number was not found as plain text.");
                    if (containsExponential)
                    {
                        Console.WriteLine("The HTML contains an exponential representation of the number.");
                    }
                    else
                    {
                        Console.WriteLine("The expected plain text representation was not found.");
                    }
                }

                // Output a snippet of the HTML for manual inspection
                Console.WriteLine();
                Console.WriteLine("HTML snippet around the cell value:");
                int index = htmlContent.IndexOf(largeNumberString, StringComparison.Ordinal);
                if (index >= 0)
                {
                    int snippetStart = Math.Max(0, index - 50);
                    int snippetLength = Math.Min(200, htmlContent.Length - snippetStart);
                    Console.WriteLine(htmlContent.Substring(snippetStart, snippetLength));
                }
                else
                {
                    Console.WriteLine("(Number not found in HTML content)");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
