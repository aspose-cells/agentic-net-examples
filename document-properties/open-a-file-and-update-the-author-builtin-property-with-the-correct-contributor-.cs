using System;
using Aspose.Cells;

namespace UpdateAuthorProperty
{
    class Program
    {
        static void Main()
        {
            // Path to the existing workbook
            string inputPath = "InputWorkbook.xlsx";

            // Load the workbook from the file system
            Workbook workbook = new Workbook(inputPath);

            // Set the built‑in Author property to the desired contributor identifier
            string contributorId = "Contributor_12345";
            workbook.BuiltInDocumentProperties.Author = contributorId;

            // Save the workbook with the updated property
            string outputPath = "UpdatedWorkbook.xlsx";
            workbook.Save(outputPath);

            Console.WriteLine($"Author property updated to '{contributorId}' and saved to '{outputPath}'.");
        }
    }
}