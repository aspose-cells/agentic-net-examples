// Title: Load HTML with an <img> tag into an Aspose.Cells workbook in C# while ensuring images are ignored and only text is imported
// AI Prompts: Use Aspose.Cells HtmlLoadOptions to load an HTML string containing an <img> element into a Workbook and confirm that the image is not imported into any cell. | Retrieve the HtmlString of the target cell and verify that it does not contain an <img> tag, then enumerate cells with place‑in‑cell pictures to ensure the count is zero. | After validation, save the workbook as an XLSX file and output the cell value to demonstrate that only the textual content was preserved.
// Common Searches: c# aspose.cells load html string without embedding images | how to prevent img tags from creating picture cells when importing html in aspose.cells | verify that GetCellsWithPlaceInCellPicture returns zero after html import aspose.cells | check cell HtmlString for image tags using aspose.cells c# | save workbook after html import ensuring only text is kept aspose.cells
// Tags: htmlloadoptions ignore images aspose.cells | validate cell htmlstring no img tag c# | enumerate place-in-cell pictures aspose.cells | save workbook xlsx after html import c# | import html string to workbook aspose.cells

using System;
using System.IO;
using System.Text;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsHtmlImportDemo
{
    // The example loads an HTML string that includes an <img> tag into an Aspose.Cells Workbook using HtmlLoadOptions, confirms that the image is ignored, validates that the cell's HtmlString contains only text, checks that no cells have place‑in‑cell pictures, and saves the result as an XLSX file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Sample HTML containing text and an <img> tag.
                string htmlContent = "<p>Hello World<img src='image.png' alt='test'></p>";

                // Convert the HTML string to a memory stream.
                byte[] htmlBytes = Encoding.UTF8.GetBytes(htmlContent);
                using (MemoryStream htmlStream = new MemoryStream(htmlBytes))
                {
                    // Load options for HTML. No special options needed; images are ignored.
                    HtmlLoadOptions loadOptions = new HtmlLoadOptions(LoadFormat.Html);

                    // Load the HTML content into a workbook.
                    Workbook workbook = new Workbook(htmlStream, loadOptions);

                    // Access the first worksheet and the first cell where the text is placed.
                    Worksheet worksheet = workbook.Worksheets[0];
                    Cell cell = worksheet.Cells["A1"];

                    // Verify that the cell contains only the textual part ("Hello World").
                    Console.WriteLine("Cell A1 value: " + cell.StringValue);
                    // Expected output: "Hello World"

                    // Ensure that the HTML string stored in the cell does not contain the <img> tag.
                    string cellHtml = cell.HtmlString;
                    bool imgTagPresent = cellHtml != null &&
                                         cellHtml.IndexOf("<img", StringComparison.OrdinalIgnoreCase) >= 0;
                    Console.WriteLine("Image tag present in cell HTML: " + imgTagPresent);
                    // Expected output: False

                    // Verify that no cells contain embedded pictures (image tags are ignored).
                    int pictureCellCount = 0;
                    IEnumerator enumerator = worksheet.Cells.GetCellsWithPlaceInCellPicture();
                    while (enumerator.MoveNext())
                    {
                        pictureCellCount++;
                    }
                    Console.WriteLine("Number of cells with embedded pictures: " + pictureCellCount);
                    // Expected output: 0

                    // Save the workbook (ensure the directory exists).
                    string outputPath = "HtmlImportResult.xlsx";
                    try
                    {
                        workbook.Save(outputPath, SaveFormat.Xlsx);
                        Console.WriteLine($"Workbook saved to '{outputPath}'.");
                    }
                    catch (Exception saveEx)
                    {
                        Console.WriteLine("Error saving workbook: " + saveEx.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
