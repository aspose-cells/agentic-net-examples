using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsHtmlHeadingVerification
{
    class Program
    {
        static void Main()
        {
            // -------------------- Create workbook --------------------
            Workbook workbook = new Workbook();

            // First worksheet with custom name
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Sales";
            sheet1.Cells["A1"].PutValue("Product");
            sheet1.Cells["B1"].PutValue("Quantity");
            sheet1.Cells["A2"].PutValue("Apple");
            sheet1.Cells["B2"].PutValue(150);

            // Second worksheet
            Worksheet sheet2 = workbook.Worksheets.Add("Inventory");
            sheet2.Cells["A1"].PutValue("Item");
            sheet2.Cells["B1"].PutValue("Stock");
            sheet2.Cells["A2"].PutValue("Banana");
            sheet2.Cells["B2"].PutValue(300);

            // -------------------- Save as single HTML file --------------------
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                SaveAsSingleFile = true,   // all sheets in one HTML file
                ShowAllSheets = true,      // ensure all sheets are rendered
                ExportHeadings = true      // export row/column headings (optional)
            };
            string htmlPath = "WorkbookExport.html";
            workbook.Save(htmlPath, saveOptions);

            // -------------------- Load generated HTML --------------------
            string htmlContent = File.ReadAllText(htmlPath);

            // -------------------- Verify heading tags contain worksheet titles --------------------
            bool salesHeadingFound = Regex.IsMatch(htmlContent, @"<h[1-6][^>]*>\s*Sales\s*</h[1-6]>", RegexOptions.IgnoreCase);
            bool inventoryHeadingFound = Regex.IsMatch(htmlContent, @"<h[1-6][^>]*>\s*Inventory\s*</h[1-6]>", RegexOptions.IgnoreCase);

            Console.WriteLine($"Heading for 'Sales' sheet found: {salesHeadingFound}");
            Console.WriteLine($"Heading for 'Inventory' sheet found: {inventoryHeadingFound}");

            // Simple assertion (could be replaced with a testing framework)
            if (salesHeadingFound && inventoryHeadingFound)
            {
                Console.WriteLine("HTML verification succeeded: all worksheet titles are present as headings.");
            }
            else
            {
                Console.WriteLine("HTML verification failed: missing worksheet title headings.");
            }
        }
    }
}