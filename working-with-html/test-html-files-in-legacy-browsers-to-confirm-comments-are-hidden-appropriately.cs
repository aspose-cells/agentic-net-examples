// Title: Verify that hidden cell comments are excluded when exporting an Aspose.Cells workbook to HTML using C#
// AI Prompts: Create a C# console application that generates a workbook, adds a comment to a cell, sets Comment.IsVisible = false, saves the workbook as HTML with HtmlSaveOptions, reads the resulting HTML file, and confirms the comment text is absent. | Adapt existing Aspose.Cells code to ensure hidden comments are omitted from the HTML output and output a pass/fail result based on whether the comment string appears in the file.
// Common Searches: how to prevent Aspose.Cells hidden comments from appearing in exported HTML C# | C# Aspose.Cells export to HTML without cell comments | verify hidden comment removal in Aspose.Cells HTML output | Aspose.Cells HtmlSaveOptions hide comments for legacy browsers | check generated HTML for absent comment text using Aspose.Cells C#
// Tags: Aspose.Cells HTML export without cell comments | C# HtmlSaveOptions comment visibility control | detect absent comment text in generated HTML | legacy browser compatibility Aspose.Cells HTML output | cell comment IsVisible false Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsLegacyHtmlTest
{
    // The sample creates a workbook, adds a hidden comment to cell A1, saves the workbook as HTML using HtmlSaveOptions, reads the HTML file, and verifies that the comment text does not appear, printing PASS or FAIL.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Paths for temporary files
                string excelPath = "TestWorkbook.xlsx";
                string htmlPath = "TestWorkbook.html";

                // 1. Create a new workbook and add a comment to a cell
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Data";

                // Sample data
                sheet.Cells["A1"].PutValue("Hello");
                sheet.Cells["B1"].PutValue("World");

                // Add a comment to cell A1 and hide it
                int commentIdx = sheet.Comments.Add("A1");
                Comment comment = sheet.Comments[commentIdx];
                comment.Note = "This is a test comment";
                comment.IsVisible = false; // Hide comment in output

                // Save the workbook as an Excel file (optional)
                workbook.Save(excelPath);

                // 2. Save the workbook as HTML (comments are hidden because they are not visible)
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
                workbook.Save(htmlPath, htmlOptions);

                // 3. Load the generated HTML file safely
                if (!File.Exists(htmlPath))
                {
                    Console.WriteLine("Error: HTML file was not generated.");
                    return;
                }

                string htmlContent;
                try
                {
                    htmlContent = File.ReadAllText(htmlPath);
                }
                catch (Exception readEx)
                {
                    Console.WriteLine($"Error reading HTML file: {readEx.Message}");
                    return;
                }

                // 4. Verify that the comment text does not appear in the HTML output
                string commentText = "This is a test comment";
                bool commentFound = htmlContent.Contains(commentText);

                // 5. Output the verification result
                Console.WriteLine(commentFound
                    ? "FAIL: Comment text was found in the HTML output."
                    : "PASS: Comment text is hidden in the HTML output.");

                // Optional cleanup
                // if (File.Exists(excelPath)) File.Delete(excelPath);
                // if (File.Exists(htmlPath)) File.Delete(htmlPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Exception: {ex.Message}");
            }
        }
    }
}
