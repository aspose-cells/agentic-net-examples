using System;
using System.IO;
using Aspose.Cells;

class ValidateExternalLinkPath
{
    static void Main()
    {
        // Load the workbook (lifecycle rule: load)
        string sourceFile = "input.xlsx";
        Workbook workbook = new Workbook(sourceFile);

        // Define the new external link path
        string newExternalPath = @"C:\Data\external.xlsx";

        // Validate that the file exists before applying the change
        if (File.Exists(newExternalPath))
        {
            // Update each external link's DataSource (feature rule: ExternalLink.DataSource)
            foreach (ExternalLink link in workbook.Worksheets.ExternalLinks)
            {
                link.DataSource = newExternalPath;
            }

            // Save the workbook (lifecycle rule: save)
            workbook.Save("output.xlsx");
            Console.WriteLine("External link path updated and workbook saved.");
        }
        else
        {
            Console.WriteLine($"The specified external file does not exist: {newExternalPath}");
            Console.WriteLine("No changes were applied to the workbook.");
        }
    }
}