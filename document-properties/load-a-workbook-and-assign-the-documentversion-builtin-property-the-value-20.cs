using System;
using Aspose.Cells;

namespace AsposeCellsDocumentVersionDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook from disk
            string inputPath = "input.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Set the built‑in DocumentVersion property to "2.0"
            workbook.BuiltInDocumentProperties.DocumentVersion = "2.0";

            // Save the workbook with the updated property
            string outputPath = "output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"DocumentVersion set to \"2.0\" and workbook saved to {outputPath}");
        }
    }
}