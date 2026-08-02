using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsHtmlSpaceDemo
{
    // Author: Aspose.Cells .NET example – demonstrates impact of removing redundant spaces on HTML processing performance
    class Program
    {
        static void Main()
        {
            // Sample HTML containing many redundant spaces
            string html = "<p>   This    text   has   redundant   spaces   </p>";

            // Convert the HTML string to a memory stream
            byte[] htmlBytes = Encoding.UTF8.GetBytes(html);
            using (MemoryStream htmlStream = new MemoryStream(htmlBytes))
            {
                // ---------- Load HTML with space removal ----------
                HtmlLoadOptions loadOpts = new HtmlLoadOptions
                {
                    // Deleting redundant spaces reduces the amount of text stored in cells,
                    // which speeds up formula calculation and subsequent saving.
                    DeleteRedundantSpaces = true
                };

                Stopwatch swLoad = Stopwatch.StartNew();
                Workbook wb = new Workbook(htmlStream, loadOpts);
                swLoad.Stop();

                // Access the loaded cell to verify space removal
                string cellText = wb.Worksheets[0].Cells["A1"].StringValue;
                Console.WriteLine($"Cell text after loading (spaces removed): \"{cellText}\"");

                // ---------- Save to HTML ----------
                HtmlSaveOptions saveOpts = new HtmlSaveOptions
                {
                    // Excluding unused styles further reduces file size and improves rendering speed.
                    ExcludeUnusedStyles = true
                };

                Stopwatch swSave = Stopwatch.StartNew();
                wb.Save("output.html", saveOpts);
                swSave.Stop();

                // Output performance measurements
                Console.WriteLine($"Load time with DeleteRedundantSpaces=true: {swLoad.ElapsedMilliseconds} ms");
                Console.WriteLine($"Save time with ExcludeUnusedStyles=true: {swSave.ElapsedMilliseconds} ms");
            }
        }
    }
}