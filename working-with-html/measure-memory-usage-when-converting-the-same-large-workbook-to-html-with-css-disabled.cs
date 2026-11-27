// Title: Measure private memory consumption when exporting a large workbook to HTML with CSS disabled using Aspose.Cells for .NET
// Common Searches: how to measure memory usage of Aspose.Cells HTML export in C# | C# Aspose.Cells export large workbook to HTML without external CSS and track memory | profile private memory before and after workbook.Save as HTML using Aspose.Cells | benchmark Aspose.Cells HTML conversion memory footprint in .NET | disable CSS export in HtmlSaveOptions and monitor process memory with Aspose.Cells
// Tags: Aspose.Cells HTML export memory profiling | disable CSS in HtmlSaveOptions Aspose.Cells | large workbook to HTML performance Aspose.Cells | process private memory measurement C# | embed images as base64 Aspose.Cells HTML

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;
using System.Drawing;

namespace AsposeCellsMemoryMeasurement
{
    // The example creates a 5,000‑row by 50‑column workbook, adds random numeric data, occasional text, formulas, and header styling, then saves it to HTML with ExportWorksheetCSSSeparately set to false and images embedded as base64. It records the process's private memory size before and after the Save operation, forces garbage collection, and outputs the memory usage delta.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Path for the generated HTML file
                string outputHtmlPath = Path.Combine(Path.GetTempPath(), "LargeWorkbook_NoCss.html");

                // Create a large workbook programmatically
                Workbook workbook = CreateLargeWorkbook();

                // Prepare HTML save options with CSS disabled
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    // Disable exporting CSS to a separate file (styles will be inline)
                    ExportWorksheetCSSSeparately = false,
                    // Embed images as base64 to avoid external resources
                    ExportImagesAsBase64 = true
                };

                // Measure memory usage during the conversion
                long memoryBefore = GetCurrentPrivateMemory();
                Console.WriteLine($"Memory before conversion: {FormatBytes(memoryBefore)}");

                // Perform the conversion
                workbook.Save(outputHtmlPath, htmlOptions);

                // Force garbage collection to get a more accurate post‑conversion measurement
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();

                long memoryAfter = GetCurrentPrivateMemory();
                Console.WriteLine($"Memory after conversion:  {FormatBytes(memoryAfter)}");

                long memoryDelta = memoryAfter - memoryBefore;
                Console.WriteLine($"Memory increase:          {FormatBytes(memoryDelta)}");

                Console.WriteLine($"HTML file saved to: {outputHtmlPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        /// <returns>A populated Workbook instance.</returns>
        private static Workbook CreateLargeWorkbook()
        {
            const int totalRows = 5000;   // Adjust as needed for your environment
            const int totalCols = 50;

            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            Cells cells = ws.Cells;

            Random rnd = new Random();

            for (int row = 0; row < totalRows; row++)
            {
                for (int col = 0; col < totalCols; col++)
                {
                    // Populate with random numbers and occasional strings
                    if (rnd.NextDouble() < 0.1)
                        cells[row, col].PutValue($"Text {row}-{col}");
                    else
                        cells[row, col].PutValue(rnd.NextDouble() * 10000);
                }

                // Occasionally add a formula to increase complexity
                if (row % 100 == 0 && totalCols > 2)
                {
                    string formula = $"=SUM(A{row + 1}:B{row + 1})";
                    cells[row, totalCols - 1].Formula = formula;
                }
            }

            // Apply some formatting to increase the rendering workload
            Style style = wb.CreateStyle();
            style.Font.Size = 10;
            style.Font.Color = Color.DarkBlue;
            style.Font.IsBold = true;
            style.ForegroundColor = Color.LightYellow;
            style.Pattern = BackgroundType.Solid;

            // Apply style to the header row
            for (int col = 0; col < totalCols; col++)
            {
                cells[0, col].SetStyle(style);
            }

            return wb;
        }

        private static long GetCurrentPrivateMemory()
        {
            using (Process proc = Process.GetCurrentProcess())
            {
                proc.Refresh();
                return proc.PrivateMemorySize64;
            }
        }

        /// <summary>
        /// Formats a byte count into a human‑readable string.
        /// </summary>
        private static string FormatBytes(long bytes)
        {
            const double KB = 1024.0;
            const double MB = KB * 1024.0;
            const double GB = MB * 1024.0;

            if (bytes >= GB)
                return $"{bytes / GB:F2} GB";
            if (bytes >= MB)
                return $"{bytes / MB:F2} MB";
            if (bytes >= KB)
                return $"{bytes / KB:F2} KB";

            return $"{bytes} B";
        }
    }
}
