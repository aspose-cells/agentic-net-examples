// Title: Create an Excel workbook with Aspose.Cells, export to HTML, and verify each <td> has a data‑cell attribute in C#
// AI Prompts: Write a C# console program that builds a Workbook, populates sample cells, saves it as HTML to a MemoryStream using Aspose.Cells, reads the HTML string, and checks every <td> tag for a data‑cell attribute, reporting any missing ones. | Enhance the verification loop to output the original Excel cell address (e.g., A2) when a <td> element does not contain the data‑cell attribute.
// Common Searches: aspocells c# export workbook to html and validate custom data-cell attribute in table cells | how to check for data-cell attribute in generated html from Aspose.Cells using regex | c# read Aspose.Cells html output from memory stream and verify td attributes | detect missing data-cell attributes in html tables produced by Aspose.Cells | sample code for verifying html cells after saving Excel as html with Aspose.Cells
// Tags: Aspose.Cells HTML export verification | C# regex td attribute check | MemoryStream Aspose.Cells HTML | Excel-to-HTML data-cell attribute | Automated HTML cell validation C#

using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsHtmlVerification
{
    // The example creates a workbook, fills sample data, saves it as HTML to a MemoryStream with Aspose.Cells, reads the HTML, uses a regular expression to locate all <td> tags, verifies each contains a data‑cell attribute, reports any missing attributes (including the cell's inner text), and prints a success or failure message.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // 1. Create a new workbook and add some sample data
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Header1");
                sheet.Cells["B1"].PutValue("Header2");
                sheet.Cells["A2"].PutValue("Row1Col1");
                sheet.Cells["B2"].PutValue("Row1Col2");
                sheet.Cells["A3"].PutValue("Row2Col1");
                sheet.Cells["B3"].PutValue("Row2Col2");

                // 2. Save the workbook as HTML into a memory stream
                using (MemoryStream htmlStream = new MemoryStream())
                {
                    // HtmlSaveOptions already defaults to SaveFormat.Html, no need to set SaveFormat
                    HtmlSaveOptions saveOptions = new HtmlSaveOptions();

                    workbook.Save(htmlStream, saveOptions);
                    htmlStream.Position = 0; // Reset stream position for reading

                    // 3. Read the generated HTML as a string
                    string htmlContent;
                    using (StreamReader reader = new StreamReader(htmlStream, Encoding.UTF8))
                    {
                        htmlContent = reader.ReadToEnd();
                    }

                    // 4. Verify that every <td> element has a "data-cell" attribute using regex
                    Regex tdRegex = new Regex(@"<td\b[^>]*>", RegexOptions.IgnoreCase);
                    MatchCollection tdMatches = tdRegex.Matches(htmlContent);

                    if (tdMatches.Count == 0)
                    {
                        Console.WriteLine("No <td> elements found in the generated HTML.");
                        return;
                    }

                    bool allCellsHaveAttribute = true;
                    foreach (Match match in tdMatches)
                    {
                        string tdTag = match.Value;
                        // Check if the tag contains a data-cell attribute
                        if (!Regex.IsMatch(tdTag, @"\bdata-cell\s*=", RegexOptions.IgnoreCase))
                        {
                            allCellsHaveAttribute = false;
                            // Extract inner text for reporting (simple approximation)
                            string innerTextPattern = $@"{Regex.Escape(tdTag)}(.*?)</td>";
                            Match innerMatch = Regex.Match(htmlContent, innerTextPattern, RegexOptions.Singleline | RegexOptions.IgnoreCase);
                            string innerText = innerMatch.Success ? innerMatch.Groups[1].Value.Trim() : string.Empty;
                            Console.WriteLine($"Missing data-cell attribute in cell with inner text: \"{innerText}\"");
                        }
                    }

                    if (allCellsHaveAttribute)
                    {
                        Console.WriteLine("Verification succeeded: every <td> element has a data-cell attribute.");
                    }
                    else
                    {
                        Console.WriteLine("Verification failed: some <td> elements are missing the data-cell attribute.");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
