using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExternalLinkValidation
{
    class Program
    {
        static void Main()
        {
            // Load the workbook that contains external links
            Workbook workbook = new Workbook("input.xlsx");

            // Define the new external link path you want to apply
            string newExternalPath = @"C:\Data\NewExternalSource.xlsx";

            // Validate that the new path points to an existing file
            if (!File.Exists(newExternalPath))
            {
                Console.WriteLine($"Error: The file '{newExternalPath}' does not exist. External link not updated.");
                return;
            }

            // Apply the new data source to each external link in the workbook
            foreach (ExternalLink link in workbook.Worksheets.ExternalLinks)
            {
                link.DataSource = newExternalPath;
                Console.WriteLine($"External link updated to: {link.DataSource}");
            }

            // Save the workbook with the updated external link(s)
            workbook.Save("output.xlsx");
            Console.WriteLine("Workbook saved successfully with validated external link.");
        }
    }
}