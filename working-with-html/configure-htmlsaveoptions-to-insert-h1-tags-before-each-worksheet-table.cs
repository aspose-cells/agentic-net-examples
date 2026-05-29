using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsHtmlH1Demo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add sample data to multiple worksheets
            Workbook workbook = new Workbook();

            // First worksheet
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Sales";
            sheet1.Cells["A1"].PutValue("Product");
            sheet1.Cells["B1"].PutValue("Quantity");
            sheet1.Cells["A2"].PutValue("Apple");
            sheet1.Cells["B2"].PutValue(120);

            // Second worksheet
            Worksheet sheet2 = workbook.Worksheets.Add("Inventory");
            sheet2.Cells["A1"].PutValue("Item");
            sheet2.Cells["B1"].PutValue("Stock");
            sheet2.Cells["A2"].PutValue("Banana");
            sheet2.Cells["B2"].PutValue(85);

            // Configure HTML save options
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            // (Other options can be set here if needed)

            // Save the workbook to a memory stream as HTML
            using (MemoryStream htmlStream = new MemoryStream())
            {
                workbook.Save(htmlStream, saveOptions);
                htmlStream.Position = 0;

                // Read the generated HTML into a string
                string htmlContent = new StreamReader(htmlStream, Encoding.UTF8).ReadToEnd();

                // Insert <h1> tags before each <table> element.
                // The order of tables corresponds to the order of worksheets.
                int sheetIndex = 0;
                string modifiedHtml = Regex.Replace(
                    htmlContent,
                    @"<table",
                    match =>
                    {
                        // Get the worksheet name for the current table
                        string sheetName = workbook.Worksheets[sheetIndex].Name;
                        sheetIndex++;

                        // Insert an <h1> tag before the table tag
                        return $"<h1>{System.Web.HttpUtility.HtmlEncode(sheetName)}</h1>\n<table";
                    },
                    RegexOptions.IgnoreCase);

                // Write the modified HTML to a file
                string outputPath = "WorkbookWithHeadings.html";
                File.WriteAllText(outputPath, modifiedHtml, Encoding.UTF8);

                Console.WriteLine($"HTML file saved with <h1> headings: {outputPath}");
            }
        }
    }
}