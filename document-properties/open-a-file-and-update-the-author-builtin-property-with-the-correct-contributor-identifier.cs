using System;
using Aspose.Cells;

namespace UpdateAuthorProperty
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the existing workbook
            string inputPath = "input.xlsx";

            // Load the workbook from the file
            Workbook workbook = new Workbook(inputPath);

            // Update the built‑in Author property with the contributor identifier
            string contributorId = "Contributor_12345";
            workbook.BuiltInDocumentProperties.Author = contributorId;

            // Save the workbook back (overwrites the original file or specify a new path)
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);

            Console.WriteLine($"Author property updated to '{contributorId}' and saved to '{outputPath}'.");
        }
    }
}