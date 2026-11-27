// Title: Compare Aspose.Cells HTML output with PresentationPreference BestFit versus default layout in C#
// Common Searches: Aspose.Cells C# generate HTML with PresentationPreference true and compare to default | difference between BestFit and default HTML layout using Aspose.Cells HtmlSaveOptions | measure HTML file size and row count when using PresentationPreference in Aspose.Cells | compare visual fidelity of Aspose.Cells HTML output with and without BestFit | C# code to save workbook as HTML with PresentationPreference = true
// Tags: Aspose.Cells HtmlSaveOptions PresentationPreference BestFit | HTML output size analysis Aspose.Cells | auto‑fit columns vs BestFit layout Aspose.Cells | HTML table row count comparison C# Aspose.Cells | visual fidelity assessment Aspose.Cells HTML

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsHtmlPresentationComparison
{
    // The sample creates a workbook, adds data, merges a long text cell, applies header styling, and auto‑fits columns. It then saves the workbook twice as HTML: once with the default layout (PresentationPreference = false) and once with the BestFit layout (PresentationPreference = true). After saving, the program reports each file's byte size and counts the <tr> rows in the generated HTML, highlighting any structural or visual differences between the two layouts.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // -----------------------------------------------------------------
                // 1. Create a sample workbook with diverse content to highlight layout differences
                // -----------------------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate some data
                cells["A1"].PutValue("Product");
                cells["B1"].PutValue("Price");
                cells["C1"].PutValue("Description");

                cells["A2"].PutValue("Apple");
                cells["B2"].PutValue(1.25);
                cells["C2"].PutValue("Fresh red apples from the orchard.");

                cells["A3"].PutValue("Banana");
                cells["B3"].PutValue(0.75);
                cells["C3"].PutValue("Ripe yellow bananas, high in potassium.");

                // Add a merged cell with longer text to test auto‑fit behavior
                cells.Merge(5, 0, 1, 5); // Merge F6:J6 (zero‑based indices)
                cells[5, 0].PutValue("This is a long merged cell that should wrap or expand depending on the presentation preference setting.");

                // Apply some formatting
                Style headerStyle = workbook.CreateStyle();
                headerStyle.Font.IsBold = true;
                headerStyle.ForegroundColor = System.Drawing.Color.LightGray;
                headerStyle.Pattern = BackgroundType.Solid;
                cells["A1"].SetStyle(headerStyle);
                cells["B1"].SetStyle(headerStyle);
                cells["C1"].SetStyle(headerStyle);

                // Auto‑fit columns (default layout will be affected by PresentationPreference later)
                sheet.AutoFitColumns();

                // -----------------------------------------------------------------
                // 2. Define output folder and file names
                // -----------------------------------------------------------------
                string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "HtmlOutput");
                Directory.CreateDirectory(outputDir);

                string defaultHtmlPath = Path.Combine(outputDir, "DefaultLayout.html");
                string bestFitHtmlPath = Path.Combine(outputDir, "BestFitLayout.html");

                // -----------------------------------------------------------------
                // 3. Save using default HtmlSaveOptions (PresentationPreference = false)
                // -----------------------------------------------------------------
                HtmlSaveOptions defaultOptions = new HtmlSaveOptions(SaveFormat.Html)
                {
                    PresentationPreference = false, // default layout
                    Encoding = Encoding.UTF8
                };

                workbook.Save(defaultHtmlPath, defaultOptions);
                Console.WriteLine($"Saved default layout HTML to: {defaultHtmlPath}");

                // -----------------------------------------------------------------
                // 4. Save using PresentationPreference = true (BestFit)
                // -----------------------------------------------------------------
                HtmlSaveOptions bestFitOptions = new HtmlSaveOptions(SaveFormat.Html)
                {
                    PresentationPreference = true, // enable BestFit layout
                    Encoding = Encoding.UTF8
                };

                workbook.Save(bestFitHtmlPath, bestFitOptions);
                Console.WriteLine($"Saved BestFit layout HTML to: {bestFitHtmlPath}");

                // -----------------------------------------------------------------
                // 5. Compare the two generated HTML files
                // -----------------------------------------------------------------
                FileInfo defaultInfo = new FileInfo(defaultHtmlPath);
                FileInfo bestFitInfo = new FileInfo(bestFitHtmlPath);

                Console.WriteLine();
                Console.WriteLine("=== Comparison Summary ===");
                Console.WriteLine($"Default layout file size : {defaultInfo.Length} bytes");
                Console.WriteLine($"BestFit layout file size : {bestFitInfo.Length} bytes");
                Console.WriteLine($"Size difference          : {Math.Abs(bestFitInfo.Length - defaultInfo.Length)} bytes");

                // Simple visual fidelity hint: count number of <tr> rows generated
                int defaultRowCount = CountHtmlTableRows(defaultHtmlPath);
                int bestFitRowCount = CountHtmlTableRows(bestFitHtmlPath);

                Console.WriteLine($"Default layout table rows: {defaultRowCount}");
                Console.WriteLine($"BestFit layout table rows: {bestFitRowCount}");
                Console.WriteLine();

                // Provide a textual recommendation based on observed differences
                if (bestFitRowCount != defaultRowCount || bestFitInfo.Length != defaultInfo.Length)
                {
                    Console.WriteLine("The BestFit layout produced a different HTML structure, which usually results in a more compact and visually appealing rendering.");
                }
                else
                {
                    Console.WriteLine("Both layouts generated identical HTML structures. In this simple scenario, the visual difference may be minimal.");
                }

                Console.WriteLine();
                Console.WriteLine("Open the generated HTML files in a browser to manually verify visual fidelity.");
            }
            catch (Exception ex)
            {
                // Output full exception details for debugging
                Console.WriteLine("An error occurred during processing:");
                Console.WriteLine(ex.ToString());
            }
        }

        private static int CountHtmlTableRows(string htmlFilePath)
        {
            if (!File.Exists(htmlFilePath))
                return 0;

            int count = 0;
            try
            {
                using (StreamReader reader = new StreamReader(htmlFilePath))
                {
                    string line;
                    while ((line = reader.ReadLine()) != null)
                    {
                        int index = 0;
                        while ((index = line.IndexOf("<tr", index, StringComparison.OrdinalIgnoreCase)) >= 0)
                        {
                            count++;
                            index += 3; // Move past the found tag
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to read '{htmlFilePath}': {ex.Message}");
            }
            return count;
        }
    }
}
