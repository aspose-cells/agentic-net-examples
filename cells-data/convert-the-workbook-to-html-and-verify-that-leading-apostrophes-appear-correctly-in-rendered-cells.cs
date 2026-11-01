// Title: Export an Excel workbook to HTML with Aspose.Cells and confirm that leading apostrophes are retained in cell output
// AI Prompts: Generate C# code that loads or creates an Excel file, writes a string beginning with an apostrophe into a cell, saves the workbook as HTML using Aspose.Cells HtmlSaveOptions, and reads the resulting HTML to verify the apostrophe is present. | Show how to inspect the generated HTML for the encoded apostrophe (&#39;) after converting a worksheet that contains a leading apostrophe with Aspose.Cells.
// Common Searches: Aspose.Cells C# export worksheet to HTML keep leading apostrophe | how to test apostrophe rendering in HTML output from Excel using Aspose.Cells | verify special characters are not stripped when saving Excel as HTML with Aspose.Cells | HtmlSaveOptions preserve single quote at start of cell value Aspose.Cells | C# read generated HTML file to confirm &#39; appears after Excel to HTML conversion
// Tags: Aspose.Cells HTML export leading apostrophe | C# HtmlSaveOptions special character handling | Excel to HTML conversion apostrophe preservation | validate generated HTML for encoded quotes Aspose.Cells | cell value starting with apostrophe export

using System;
using System.IO;
using Aspose.Cells;

// The program loads or creates a workbook, inserts a value with a leading apostrophe into cell A1, saves the workbook as HTML using Aspose.Cells HtmlSaveOptions, reads the HTML file, and checks for the encoded apostrophe (&#39;) to ensure it is rendered correctly.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one.
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook(); // creates a default workbook with one worksheet
            }

            // Ensure a cell contains a leading apostrophe that should be displayed.
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("''Hello");

            // Convert the workbook to HTML.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            // No special option needed; Aspose.Cells renders the cell value as it appears.
            workbook.Save(outputPath, htmlOptions);

            // Verify that the generated HTML contains the leading apostrophe.
            if (File.Exists(outputPath))
            {
                string htmlContent = File.ReadAllText(outputPath);

                // Aspose.Cells HTML‑encodes the apostrophe as &#39;.
                if (htmlContent.Contains("&#39;Hello") || htmlContent.Contains("'Hello"))
                {
                    Console.WriteLine("Leading apostrophe rendered correctly in HTML.");
                }
                else
                {
                    Console.WriteLine("Leading apostrophe is missing in the HTML output.");
                }
            }
            else
            {
                Console.WriteLine($"Failed to generate HTML file: {outputPath}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
