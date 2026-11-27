// Title: Generate HTML with all workbook styles and then a reduced‑style version using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export workbook to HTML preserving every style | How to disable ExcludeUnusedStyles when saving Excel as HTML with Aspose.Cells | Remove unused styles from a workbook before HTML conversion using Aspose.Cells | Compare full‑style and minimized‑style HTML output from the same Excel file in C# | Aspose.Cells SaveFormat.Html example with ExcludeUnusedStyles false
// Tags: HtmlSaveOptions ExcludeUnusedStyles false | RemoveUnusedStyles workbook Aspose.Cells | full style HTML export Aspose.Cells | reduced style HTML export Aspose.Cells | C# compare HTML style outputs Excel

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsHtmlStyleComparison
{
    // The sample loads or creates an Excel workbook, saves it to HTML with ExcludeUnusedStyles set to false to retain every style, then calls RemoveUnusedStyles on the workbook and saves a second HTML file using the default ExcludeUnusedStyles (true). This produces a full‑style HTML file and a reduced‑style HTML file for side‑by‑side comparison.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputWorkbookPath = "input.xlsx";
            const string fullStyleHtmlPath = "full_styles.html";
            const string reducedStyleHtmlPath = "reduced_styles.html";

            try
            {
                // -----------------------------------------------------------------
                // 1. Load an existing workbook or create a new one with sample data.
                // -----------------------------------------------------------------
                Workbook workbook;
                if (File.Exists(inputWorkbookPath))
                {
                    Console.WriteLine($"Loading workbook from \"{inputWorkbookPath}\"...");
                    workbook = new Workbook(inputWorkbookPath);
                }
                else
                {
                    Console.WriteLine("Input workbook not found. Creating a sample workbook...");
                    workbook = CreateSampleWorkbook();
                    workbook.Save(inputWorkbookPath, SaveFormat.Xlsx);
                    Console.WriteLine($"Sample workbook saved to \"{inputWorkbookPath}\".");
                }

                // ---------------------------------------------------------------
                // 2. Save HTML with all styles (ExcludeUnusedStyles = false)
                // ---------------------------------------------------------------
                var fullStyleOptions = new HtmlSaveOptions(SaveFormat.Html)
                {
                    ExcludeUnusedStyles = false,               // keep every style in the output
                    ExportActiveWorksheetOnly = false          // export whole workbook
                };
                Console.WriteLine($"Saving HTML with full style set to \"{fullStyleHtmlPath}\"...");
                workbook.Save(fullStyleHtmlPath, fullStyleOptions);
                Console.WriteLine("Full‑style HTML saved successfully.");

                // ---------------------------------------------------------------
                // 3. Remove unused styles from the workbook
                // ---------------------------------------------------------------
                Console.WriteLine("Removing unused styles from the workbook...");
                workbook.RemoveUnusedStyles();

                // ---------------------------------------------------------------
                // 4. Save HTML after style reduction (default ExcludeUnusedStyles = true)
                // ---------------------------------------------------------------
                var reducedStyleOptions = new HtmlSaveOptions(SaveFormat.Html)
                {
                    ExportActiveWorksheetOnly = false // use default ExcludeUnusedStyles = true
                };
                Console.WriteLine($"Saving HTML with reduced style set to \"{reducedStyleHtmlPath}\"...");
                workbook.Save(reducedStyleHtmlPath, reducedStyleOptions);
                Console.WriteLine("Reduced‑style HTML saved successfully.");

                Console.WriteLine("Comparison files are ready:");
                Console.WriteLine($"  Full style HTML   : {fullStyleHtmlPath}");
                Console.WriteLine($"  Reduced style HTML: {reducedStyleHtmlPath}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("An error occurred during processing:");
                Console.Error.WriteLine(ex.Message);
                Console.Error.WriteLine(ex.StackTrace);
            }
        }

        private static Workbook CreateSampleWorkbook()
        {
            var wb = new Workbook();

            // Worksheet 1 – simple table with varied styles
            var ws1 = wb.Worksheets[0];
            ws1.Name = "Products";

            // Header row
            ws1.Cells["A1"].PutValue("Product");
            ws1.Cells["B1"].PutValue("Category");
            ws1.Cells["C1"].PutValue("Price");
            // Use CreateRange to obtain a Range object
            ApplyHeaderStyle(wb, ws1.Cells.CreateRange("A1:C1"));

            // Data rows with different styles
            string[,] data = {
                { "Apple", "Fruit", "1.20" },
                { "Carrot", "Vegetable", "0.80" },
                { "Steak", "Meat", "5.50" },
                { "Milk", "Dairy", "1.00" }
            };

            for (int i = 0; i < data.GetLength(0); i++)
            {
                int row = i + 2;
                ws1.Cells[$"A{row}"].PutValue(data[i, 0]);
                ws1.Cells[$"B{row}"].PutValue(data[i, 1]);
                ws1.Cells[$"C{row}"].PutValue(data[i, 2]);

                var style = wb.CreateStyle();
                style.Font.Name = "Calibri";
                style.Font.Size = 11 + i;                     // different font size per row
                style.Font.IsBold = (i % 2 == 0);
                style.ForegroundColor = (i % 2 == 0) ? Color.LightYellow : Color.LightCyan;
                style.Pattern = BackgroundType.Solid;
                ws1.Cells[$"A{row}:C{row}"].SetStyle(style);
            }

            // Worksheet 2 – contains some cells that will be deleted later,
            // leaving behind unused styles.
            var ws2 = wb.Worksheets.Add("TempData");
            ws2.Cells["A1"].PutValue("Temporary");
            var tempStyle = wb.CreateStyle();
            tempStyle.Font.Color = Color.Red;
            tempStyle.Font.IsItalic = true;
            ws2.Cells["A1"].SetStyle(tempStyle);

            // Delete the row to make the style unused
            ws2.Cells.DeleteRows(0, 1);

            return wb;
        }

        private static void ApplyHeaderStyle(Workbook wb, Aspose.Cells.Range range)
        {
            var headerStyle = wb.CreateStyle();
            headerStyle.Font.IsBold = true;
            headerStyle.Font.Color = Color.White;
            headerStyle.ForegroundColor = Color.DarkBlue;
            headerStyle.Pattern = BackgroundType.Solid;
            range.SetStyle(headerStyle);
        }
    }
}
