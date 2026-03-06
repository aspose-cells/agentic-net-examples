using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using Aspose.Cells;

namespace AsposeCellsConversionDemo
{
    class Program
    {
        static async Task Main(string[] args)
        {
            // URL of the source XLSX file to download
            string sourceUrl = "https://example.com/source.xlsx";

            // Local file paths
            string sourcePath = "source.xlsx";
            string outputPath = "output.xlsx";

            // Download the source file
            using (HttpClient client = new HttpClient())
            using (HttpResponseMessage response = await client.GetAsync(sourceUrl))
            using (Stream contentStream = await response.Content.ReadAsStreamAsync())
            using (FileStream fileStream = new FileStream(sourcePath, FileMode.Create, FileAccess.Write))
            {
                await contentStream.CopyToAsync(fileStream);
            }

            // Load the downloaded workbook (uses Workbook(string) constructor)
            Workbook workbook = new Workbook(sourcePath);

            // Save the workbook as XLSX (uses Save(string, SaveFormat) overload)
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook loaded from '{sourcePath}' and saved to '{outputPath}'.");
        }
    }
}