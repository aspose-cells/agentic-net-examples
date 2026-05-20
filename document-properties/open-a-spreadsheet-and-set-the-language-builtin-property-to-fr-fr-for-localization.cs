using System;
using Aspose.Cells;

namespace AsposeCellsLanguageDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the existing spreadsheet
            string inputPath = "input.xlsx";

            // Load the workbook (uses Aspose.Cells default load options)
            Workbook workbook = new Workbook(inputPath);

            // Set the built‑in Language property to French (France)
            workbook.BuiltInDocumentProperties.Language = "fr-FR";

            // Save the modified workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            
            Console.WriteLine($"Language property set to 'fr-FR' and workbook saved to '{outputPath}'.");
        }
    }
}