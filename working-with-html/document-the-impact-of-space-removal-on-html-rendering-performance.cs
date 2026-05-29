using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using Aspose.Cells;

class HtmlSpaceRemovalPerformanceDemo
{
    static void Main()
    {
        // HTML containing many redundant spaces
        string html = "<p>   This    text   has   redundant   spaces   </p>";

        // Convert the HTML string to a memory stream
        byte[] htmlBytes = Encoding.UTF8.GetBytes(html);
        using (MemoryStream stream = new MemoryStream(htmlBytes))
        {
            // ------------------------------------------------------------
            // Load without removing redundant spaces
            // ------------------------------------------------------------
            HtmlLoadOptions loadOptsNoDelete = new HtmlLoadOptions();
            loadOptsNoDelete.DeleteRedundantSpaces = false; // keep spaces

            Stopwatch sw = Stopwatch.StartNew();
            Workbook wbNoDelete = new Workbook(stream, loadOptsNoDelete);
            sw.Stop();

            Console.WriteLine($"Load time (no space removal): {sw.ElapsedMilliseconds} ms");
            Console.WriteLine("Cell A1 value (no removal): '" + wbNoDelete.Worksheets[0].Cells["A1"].StringValue + "'");

            // Reset stream position for the second load
            stream.Position = 0;

            // ------------------------------------------------------------
            // Load with redundant spaces removed
            // ------------------------------------------------------------
            HtmlLoadOptions loadOptsDelete = new HtmlLoadOptions();
            loadOptsDelete.DeleteRedundantSpaces = true; // delete spaces

            sw.Restart();
            Workbook wbDelete = new Workbook(stream, loadOptsDelete);
            sw.Stop();

            Console.WriteLine($"Load time (space removal enabled): {sw.ElapsedMilliseconds} ms");
            Console.WriteLine("Cell A1 value (after removal): '" + wbDelete.Worksheets[0].Cells["A1"].StringValue + "'");

            // ------------------------------------------------------------
            // Save both workbooks as HTML to compare output size
            // ------------------------------------------------------------
            wbNoDelete.Save("NoDeleteSpaces.html", SaveFormat.Html);
            wbDelete.Save("DeleteSpaces.html", SaveFormat.Html);

            FileInfo fiNoDelete = new FileInfo("NoDeleteSpaces.html");
            FileInfo fiDelete = new FileInfo("DeleteSpaces.html");

            Console.WriteLine($"HTML size without deletion: {fiNoDelete.Length} bytes");
            Console.WriteLine($"HTML size with deletion: {fiDelete.Length} bytes");
        }
    }
}