using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlImportDemo
{
    class Program
    {
        static void Main()
        {
            string htmlPath = "input.html";
            var loadOptions = new HtmlLoadOptions
            {
                LoadFormulas = true
            };

            using (var workbook = new Workbook(htmlPath, loadOptions))
            {
                string outputPath = "output.xlsx";

                if (File.Exists(outputPath))
                {
                    File.Delete(outputPath);
                }

                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"HTML data imported and saved to '{outputPath}'.");
            }
        }
    }
}