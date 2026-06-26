using System;
using Aspose.Cells;

class UpdateAuthorProperty
{
    static void Main()
    {
        // Path to the existing workbook file
        string inputPath = "input.xlsx";

        // Load the workbook from the file
        Workbook workbook = new Workbook(inputPath);

        // Set the built‑in Author property to the desired contributor identifier
        workbook.BuiltInDocumentProperties.Author = "Contributor123";

        // Save the updated workbook to a new file (or overwrite the original)
        string outputPath = "output.xlsx";
        workbook.Save(outputPath);
    }
}