using System;
using System.IO;
using Aspose.Cells;

namespace TimelineExportExample
{
    class Program
    {
        static void Main()
        {
            // Path to the source Excel macro‑enabled template (XLTM)
            string sourcePath = "TemplateWithTimeline.xltm";

            // Desired output path for the converted XLSX workbook
            string destinationPath = "ExportedTimeline.xlsx";

            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: '{sourcePath}'");
                return;
            }

            // Load the XLTM workbook (auto‑detect format)
            LoadOptions loadOptions = new LoadOptions();
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // Save as XLSX
            workbook.Save(destinationPath, SaveFormat.Xlsx);

            Console.WriteLine($"Conversion completed: '{sourcePath}' → '{destinationPath}'");
        }
    }
}