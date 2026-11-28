// Title: Export Arabic right-to-left text to HTML using Aspose.Cells default HtmlSaveOptions in C#
// Common Searches: C# Aspose.Cells export Arabic text to HTML with right-to-left alignment | How to keep RTL direction when saving a workbook as HTML using Aspose.Cells | Aspose.Cells default HtmlSaveOptions generate dir="rtl" in HTML output | Save worksheet containing Arabic characters to HTML with correct text direction in .NET | Verify RTL direction markers in HTML produced by Aspose.Cells
// Tags: export rtl text to html Aspose.Cells | default HtmlSaveOptions preserve direction | c# workbook save as html rtl alignment | arabic cell style right alignment | html output rtl direction detection

using System;
using System.IO;
using Aspose.Cells;

namespace ExportRtlHtmlDemo
{
    // The example creates a workbook, inserts Arabic text into a cell, applies a right-aligned style, saves the workbook to HTML with the default HtmlSaveOptions, and checks the generated HTML for RTL direction markers such as dir="rtl" or CSS direction settings.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the output HTML file
            string outputHtmlPath = "RtlExport.html";

            try
            {
                // ------------------------------------------------------------
                // 1. Create a new workbook and get the first worksheet
                // ------------------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // ------------------------------------------------------------
                // 2. Insert some RTL text (Arabic example) into a cell
                // ------------------------------------------------------------
                // Arabic phrase: "مرحبا بالعالم" (Hello World)
                Cell rtlCell = sheet.Cells["A1"];
                rtlCell.PutValue("مرحبا بالعالم");

                // ------------------------------------------------------------
                // 3. Apply a style that aligns text to the right (common for RTL)
                // ------------------------------------------------------------
                Style rtlStyle = workbook.CreateStyle();
                rtlStyle.HorizontalAlignment = TextAlignmentType.Right; // Align to the right
                rtlCell.SetStyle(rtlStyle);

                // ------------------------------------------------------------
                // 4. Save the workbook to HTML using the default HtmlSaveOptions
                // ------------------------------------------------------------
                workbook.Save(outputHtmlPath, SaveFormat.Html);

                // ------------------------------------------------------------
                // 5. Verify that the generated HTML file exists
                // ------------------------------------------------------------
                if (File.Exists(outputHtmlPath))
                {
                    string htmlContent = File.ReadAllText(outputHtmlPath);
                    bool containsRtl = htmlContent.Contains("dir=\"rtl\"") ||
                                       htmlContent.Contains("direction: rtl") ||
                                       htmlContent.Contains("style=\"direction:rtl\"");

                    Console.WriteLine($"HTML file '{outputHtmlPath}' has been created successfully.");
                    Console.WriteLine($"RTL direction detected in HTML: {containsRtl}");
                }
                else
                {
                    Console.WriteLine($"Failed to create the HTML file at '{outputHtmlPath}'.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred during the export process:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
