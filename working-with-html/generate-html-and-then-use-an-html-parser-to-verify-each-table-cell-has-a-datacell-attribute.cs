// Title: Export an Aspose.Cells workbook to HTML and confirm every table cell has a data‑cell attribute in C#
// Common Searches: how to save an Aspose.Cells workbook as HTML and ensure each <td> includes a data-cell attribute using C# | C# code to export Excel to HTML with data-cell attributes via Aspose.Cells | verify that generated HTML from Aspose.Cells contains data-cell on all table cells | Aspose.Cells HTML output validation for td data-cell attribute in .NET applications | parse HTML produced by Aspose.Cells to check every td element has a data-cell attribute
// Tags: Aspose.Cells HTML export with data-cell attribute | C# verification of td data-cell in generated HTML | save workbook as HTML using SaveFormat.Html | parse HTML to count td elements in .NET | validate table cell attributes in Aspose.Cells output

using System;
using System.IO;
using Aspose.Cells;

// The example creates a Workbook, fills a worksheet with sample data, saves it as an HTML file using SaveFormat.Html, then reads the file to count <td> elements and those containing the data-cell attribute, reporting whether all cells include the attribute.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add some data
            var workbook = new Workbook();
            var sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Header1");
            sheet.Cells["B1"].PutValue("Header2");
            sheet.Cells["A2"].PutValue("Row1Col1");
            sheet.Cells["B2"].PutValue("Row1Col2");
            sheet.Cells["A3"].PutValue("Row2Col1");
            sheet.Cells["B3"].PutValue("Row2Col2");

            // Define the HTML output path
            string htmlPath = Path.Combine(Directory.GetCurrentDirectory(), "output.html");

            // Ensure the directory exists
            string outputDir = Path.GetDirectoryName(htmlPath);
            if (!Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the workbook as HTML
            workbook.Save(htmlPath, SaveFormat.Html);

            // Verify the generated HTML contains the data-cell attribute
            if (!File.Exists(htmlPath))
            {
                Console.WriteLine($"Failed to generate HTML file at: {htmlPath}");
                return;
            }

            string htmlContent = File.ReadAllText(htmlPath);
            // Simple check: count occurrences of data-cell attribute in <td> elements
            int dataCellCount = 0;
            int tdCount = 0;
            using (StringReader sr = new StringReader(htmlContent))
            {
                string line;
                while ((line = sr.ReadLine()) != null)
                {
                    if (line.Contains("<td"))
                    {
                        tdCount++;
                        if (line.Contains("data-cell"))
                            dataCellCount++;
                    }
                }
            }

            if (tdCount == 0)
            {
                Console.WriteLine("No <td> elements found in the generated HTML.");
                return;
            }

            if (dataCellCount == tdCount)
                Console.WriteLine("All table cells contain the data-cell attribute.");
            else
                Console.WriteLine($"Some table cells are missing the data-cell attribute. Found {dataCellCount} out of {tdCount}.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
