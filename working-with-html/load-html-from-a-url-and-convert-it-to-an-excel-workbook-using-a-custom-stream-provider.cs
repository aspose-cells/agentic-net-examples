using System;
using System.IO;
using System.Net.Http;
using Aspose.Cells;

namespace AsposeCellsHtmlToExcel
{
    public class Program
    {
        public static void Main()
        {
            try
            {
                // URL of the HTML file to be converted
                string htmlUrl = "https://example.com/sample.html";

                // Download the HTML content as a stream
                using (HttpClient client = new HttpClient())
                using (HttpResponseMessage response = client.GetAsync(htmlUrl).Result)
                {
                    response.EnsureSuccessStatusCode();

                    using (Stream htmlStream = response.Content.ReadAsStreamAsync().Result)
                    {
                        // Load options for HTML import
                        HtmlLoadOptions loadOptions = new HtmlLoadOptions();

                        // Load the HTML stream into a workbook
                        Workbook workbook = new Workbook(htmlStream, loadOptions);

                        // Define output path and ensure its directory exists
                        string outputPath = "ConvertedFromHtml.xlsx";
                        string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                        if (!Directory.Exists(outputDir))
                        {
                            Directory.CreateDirectory(outputDir);
                        }

                        // Save the workbook as an Excel file (XLSX format)
                        workbook.Save(outputPath, SaveFormat.Xlsx);

                        Console.WriteLine($"HTML has been successfully converted to Excel: {outputPath}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}