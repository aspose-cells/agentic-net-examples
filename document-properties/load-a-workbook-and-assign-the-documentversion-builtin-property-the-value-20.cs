using System;
using Aspose.Cells;

namespace AsposeCellsDocumentVersionDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the existing workbook to be loaded
            string inputPath = "input.xlsx";

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Set the built‑in DocumentVersion property to "2.0"
            workbook.BuiltInDocumentProperties.DocumentVersion = "2.0";

            // Save the modified workbook to a new file
            string outputPath = "output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"DocumentVersion set to \"2.0\" and workbook saved to {outputPath}");
        }
    }
}