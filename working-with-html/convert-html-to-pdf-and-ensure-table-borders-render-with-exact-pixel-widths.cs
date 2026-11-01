// Title: Convert HTML to PDF with exact 1‑pixel table borders using Aspose.Cells for .NET
// AI Prompts: Write C# code that inserts a <style> block enforcing 1px solid borders on table, th, and td elements, saves the modified HTML to a temporary file, loads it with Aspose.Cells Workbook, and exports it to PDF. | Show how to set up Aspose.Cells page configuration for A4 paper size with zero margins before saving the workbook as a PDF. | Provide a C# snippet that safely deletes the temporary HTML file after the PDF conversion, handling any I/O exceptions.
// Common Searches: Aspose.Cells .NET keep exact table border thickness when converting HTML to PDF | C# convert HTML file to PDF with 1px borders using Aspose.Cells | Add custom CSS to HTML before loading into Aspose.Cells workbook for PDF export | Set zero margins and A4 page size in Aspose.Cells PDF output | How to clean up temporary files after Aspose.Cells HTML to PDF conversion in C#
// Tags: HTML to PDF conversion Aspose.Cells C# | enforce 1px table borders CSS Aspose.Cells | A4 zero‑margin page setup Aspose.Cells PDF | temporary HTML file handling Aspose.Cells | load HTML with HtmlLoadOptions Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The program reads an HTML file, injects CSS that forces 1 px solid borders on tables, writes the modified content to a temporary file, loads it into an Aspose.Cells Workbook, configures A4 paper size with zero margins, saves the workbook as a PDF, and then deletes the temporary file.
class HtmlToPdfConverter
{
    static void Main()
    {
        // Paths for input HTML and output PDF
        string htmlPath = "input.html";
        string pdfPath = "output.pdf";

        // Temporary modified HTML file path
        string tempHtmlPath = Path.Combine(Path.GetDirectoryName(htmlPath) ?? string.Empty, "temp_modified.html");

        try
        {
            // Verify input HTML exists
            if (!File.Exists(htmlPath))
                throw new FileNotFoundException($"Input HTML file not found: {htmlPath}");

            // Read the original HTML content
            string htmlContent = File.ReadAllText(htmlPath);

            // CSS to enforce exact pixel border widths for tables
            string borderCss = @"
                <style>
                    table, th, td {
                        border: 1px solid #000000 !important;
                        border-collapse: collapse;
                    }
                </style>";

            // Insert CSS before </head> if present, otherwise prepend
            if (htmlContent.IndexOf("</head>", StringComparison.OrdinalIgnoreCase) >= 0)
                htmlContent = htmlContent.Replace("</head>", borderCss + "</head>", StringComparison.OrdinalIgnoreCase);
            else
                htmlContent = borderCss + htmlContent;

            // Save modified HTML to a temporary file (Aspose.Cells works with file paths)
            File.WriteAllText(tempHtmlPath, htmlContent);

            // Ensure the temporary file was created before loading
            if (!File.Exists(tempHtmlPath))
                throw new FileNotFoundException($"Temporary HTML file not created: {tempHtmlPath}");

            // Load HTML into a workbook
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();
            Workbook workbook = new Workbook(tempHtmlPath, loadOptions);

            // Verify workbook contains at least one worksheet
            if (workbook.Worksheets.Count == 0)
                throw new InvalidOperationException("The loaded workbook does not contain any worksheets.");

            // Optional: configure PDF page settings (A4 size, zero margins)
            Worksheet sheet = workbook.Worksheets[0];
            PageSetup setup = sheet.PageSetup;
            setup.PaperSize = PaperSizeType.PaperA4; // Correct enum value for A4 size
            setup.LeftMargin = 0;
            setup.RightMargin = 0;
            setup.TopMargin = 0;
            setup.BottomMargin = 0;

            // Save as PDF
            workbook.Save(pdfPath, SaveFormat.Pdf);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
        finally
        {
            // Clean up temporary file if it exists
            if (File.Exists(tempHtmlPath))
            {
                try
                {
                    File.Delete(tempHtmlPath);
                }
                catch
                {
                    // Ignore cleanup errors
                }
            }
        }
    }
}
