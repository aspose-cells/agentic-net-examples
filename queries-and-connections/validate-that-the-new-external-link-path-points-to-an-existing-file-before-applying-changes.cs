using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExternalLinkValidation
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook that contains external links
            Workbook workbook = new Workbook("InputWorkbook.xlsx");

            // Define the new external link path you want to apply
            string newExternalPath = @"C:\ExternalData\NewSource.xlsx";

            // Validate that the new path points to an existing file
            if (!File.Exists(newExternalPath))
            {
                Console.WriteLine($"Error: The file '{newExternalPath}' does not exist.");
                return;
            }

            // Apply the new path to each external link in the workbook
            foreach (ExternalLink link in workbook.Worksheets.ExternalLinks)
            {
                // Update the DataSource property only after validation
                link.DataSource = newExternalPath;
                Console.WriteLine($"External link updated to: {link.DataSource}");
            }

            // Save the workbook with the updated external links
            workbook.Save("OutputWorkbook.xlsx");
            Console.WriteLine("Workbook saved successfully with validated external links.");
        }
    }
}