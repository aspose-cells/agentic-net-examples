using System;
using Aspose.Cells;

namespace AsposeCellsAuthorUpdate
{
    class Program
    {
        static void Main()
        {
            // Path to the existing workbook
            string inputPath = "ExistingWorkbook.xlsx";

            // Load the workbook from the file
            Workbook workbook = new Workbook(inputPath);

            // Update the built‑in Author property with the contributor identifier
            string contributorId = "Contributor_12345";
            workbook.BuiltInDocumentProperties.Author = contributorId;

            // Save the workbook (overwrites the original file or specify a new file name)
            workbook.Save(inputPath);
        }
    }
}