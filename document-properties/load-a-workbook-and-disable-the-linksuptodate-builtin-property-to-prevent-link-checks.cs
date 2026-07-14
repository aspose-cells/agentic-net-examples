using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsLinksUpToDateDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source workbook
            string inputPath = "input.xlsx";

            // Load the workbook from the file
            Workbook workbook = new Workbook(inputPath);

            // Disable the LinksUpToDate built‑in property to prevent link checks
            workbook.BuiltInDocumentProperties.LinksUpToDate = false;

            // Save the modified workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);

            // Optional: inform the user
            Console.WriteLine($"Workbook loaded from '{inputPath}', LinksUpToDate set to false, and saved to '{outputPath}'.");
        }
    }
}