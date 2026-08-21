// Title: C# Benchmark: Aspose.Cells HtmlLoadOptions DeleteRedundantSpaces Impact on Load Speed
// Description: Demonstrates how to measure the performance difference when loading HTML into an Aspose.Cells workbook with HtmlLoadOptions.DeleteRedundantSpaces enabled or disabled. The sample times each load, shows the resulting cell text, and saves both workbooks for visual comparison.
// Keywords: Aspose.Cells | HtmlLoadOptions | DeleteRedundantSpaces | HTML load performance | C# benchmark | Excel conversion speed | redundant space removal | workbook load time
// Common Searches: Aspose.Cells DeleteRedundantSpaces performance test | benchmark HTML to Excel load time C# | how does DeleteRedundantSpaces affect parsing speed | measure Aspose.Cells HTML loading speed | compare HtmlLoadOptions with and without space cleanup
// Developer Intent: Evaluate load time and cell content differences when converting HTML to Excel with and without redundant‑space removal using Aspose.Cells.
// Use Cases: Determine whether enabling DeleteRedundantSpaces improves conversion speed for large HTML reports. | Validate that space cleanup does not alter the textual data in target cells. | Generate cleaned and original Excel files to compare file size and rendering behavior.
// AI Prompts: Explain the internal processing steps of HtmlLoadOptions.DeleteRedundantSpaces and why it can speed up HTML parsing in Aspose.Cells. | Create a C# loop that processes multiple HTML files, logs load times for both DeleteRedundantSpaces settings, and outputs average results. | Recommend best practices for using DeleteRedundantSpaces when converting complex HTML documents to Excel in high‑volume scenarios.

using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using Aspose.Cells;

// Demonstrates how to measure the performance difference when loading HTML into an Aspose.Cells workbook with HtmlLoadOptions.DeleteRedundantSpaces enabled or disabled. The sample times each load, shows the resulting cell text, and saves both workbooks for visual comparison.
class HtmlSpaceRemovalPerformanceDemo
{
    static void Main()
    {
        // Sample HTML containing redundant spaces.
        string html = "<p>   This    text   has   redundant   spaces   </p>";

        // Convert the HTML string to a memory stream.
        byte[] htmlBytes = Encoding.UTF8.GetBytes(html);
        using (MemoryStream stream = new MemoryStream(htmlBytes))
        {
            // -------------------------------------------------
            // Load without deleting redundant spaces.
            // -------------------------------------------------
            HtmlLoadOptions loadOptsWithout = new HtmlLoadOptions();
            loadOptsWithout.DeleteRedundantSpaces = false; // keep spaces

            Stopwatch sw = Stopwatch.StartNew();
            Workbook wbWithout = new Workbook(stream, loadOptsWithout);
            sw.Stop();
            long elapsedWithout = sw.ElapsedMilliseconds;

            // Reset the stream position for the second load.
            stream.Position = 0;

            // -------------------------------------------------
            // Load with deleting redundant spaces.
            // -------------------------------------------------
            HtmlLoadOptions loadOptsWith = new HtmlLoadOptions();
            loadOptsWith.DeleteRedundantSpaces = true; // remove spaces

            sw.Restart();
            Workbook wbWith = new Workbook(stream, loadOptsWith);
            sw.Stop();
            long elapsedWith = sw.ElapsedMilliseconds;

            // -------------------------------------------------
            // Output performance comparison and cell values.
            // -------------------------------------------------
            Console.WriteLine($"Load time without DeleteRedundantSpaces: {elapsedWithout} ms");
            Console.WriteLine($"Load time with DeleteRedundantSpaces:    {elapsedWith} ms");

            Console.WriteLine("Cell A1 text without cleanup: '" + wbWithout.Worksheets[0].Cells["A1"].StringValue + "'");
            Console.WriteLine("Cell A1 text with cleanup:    '" + wbWith.Worksheets[0].Cells["A1"].StringValue + "'");

            // -------------------------------------------------
            // Save both workbooks to illustrate the effect.
            // -------------------------------------------------
            wbWithout.Save("output_without_cleanup.xlsx");
            wbWith.Save("output_with_cleanup.xlsx");
        }
    }
}
