using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsDivPreserveDemo
{
    class Program
    {
        static void Main()
        {
            // Paths for the source and the re‑exported HTML files
            string inputHtmlPath = "input.html";
            string outputHtmlPath = "output.html";

            // ---------- Load HTML with DIV layout support ----------
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();
            loadOptions.SupportDivTag = true;               // Preserve <div> layout while loading
            Workbook workbook = new Workbook(inputHtmlPath, loadOptions);

            // ---------- Save back to HTML preserving DIV tags ----------
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.ParseHtmlTagInCell = false;         // Keep HTML tags (e.g., <div>) as‑is in cells
            workbook.Save(outputHtmlPath, saveOptions);

            // ---------- Verify that the DIV structure is unchanged ----------
            string originalHtml = File.ReadAllText(inputHtmlPath);
            string savedHtml = File.ReadAllText(outputHtmlPath);

            // Simple verification: compare the number of <div> tags (case‑insensitive)
            int originalDivCount = Regex.Matches(originalHtml, "<div", RegexOptions.IgnoreCase).Count;
            int savedDivCount = Regex.Matches(savedHtml, "<div", RegexOptions.IgnoreCase).Count;

            Console.WriteLine($"Original DIV count: {originalDivCount}");
            Console.WriteLine($"Saved DIV count: {savedDivCount}");

            if (originalDivCount == savedDivCount)
                Console.WriteLine("DIV structure preserved.");
            else
                Console.WriteLine("DIV structure changed.");
        }
    }
}