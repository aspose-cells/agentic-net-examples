// Title: Export a workbook to HTML with Aspose.Cells and keep 64‑bit integers from converting to scientific notation
// AI Prompts: Generate C# code that inserts a 64‑bit integer into a worksheet cell, applies a numeric style that forces the full value to show, and saves the workbook as HTML using Aspose.Cells. | Add logic that reads the saved HTML file and asserts the large integer appears as plain text rather than in exponential form. | Describe how to combine HtmlSaveOptions with cell styling to prevent scientific notation in the HTML output.
// Common Searches: how to prevent scientific notation when exporting Excel to HTML with Aspose.Cells | Aspose.Cells display large 64-bit integer as plain text in HTML output | C# set numeric style for cell before saving as HTML | verify large number string in generated HTML file Aspose.Cells | Aspose.Cells HTML export large integer formatting issue
// Tags: Aspose.Cells custom numeric format plain text | HTML export large integer formatting | C# cell style full integer display | verify generated HTML contains number | HTML save options for integer formatting

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example creates a Workbook, writes a 64‑bit integer to cell A1, applies a numeric style "0" to keep the full value visible, saves the workbook as HTML with HtmlSaveOptions, then reads the HTML file to confirm the exact number is present as plain text, ensuring no scientific notation appears.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define a large number for testing
            long largeNumber = 1234567890123456789L;

            // Place the large number into cell A1
            Cell cell = sheet.Cells["A1"];
            cell.PutValue(largeNumber);

            // Apply a custom numeric format to force plain text representation
            Style style = cell.GetStyle();
            style.Custom = "0"; // Displays the full integer without scientific notation
            cell.SetStyle(style);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            // The custom format already ensures the value is rendered as plain text,
            // so no additional option is required here.

            // Save the workbook as HTML
            string htmlPath = "output.html";
            workbook.Save(htmlPath, htmlOptions);

            // Verify the generated HTML
            if (File.Exists(htmlPath))
            {
                string htmlContent = File.ReadAllText(htmlPath);
                if (htmlContent.Contains(largeNumber.ToString()))
                {
                    Console.WriteLine("Verification passed: Large number is displayed as plain text.");
                }
                else
                {
                    Console.WriteLine("Verification failed: Large number is displayed in exponential format.");
                }
            }
            else
            {
                Console.WriteLine($"Error: The file '{htmlPath}' was not created.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
