using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the existing workbook (replace with your actual file path)
        string inputFile = "input.xls";
        Workbook workbook = new Workbook(inputFile);

        // Iterate through all external links in the workbook
        for (int i = 0; i < workbook.Worksheets.ExternalLinks.Count; i++)
        {
            // Get the external link object
            ExternalLink link = workbook.Worksheets.ExternalLinks[i];

            // Retrieve the stored original data source path
            string originalPath = link.OriginalDataSource;

            // Example conversion:
            // Change a local folder path (e.g., "C:\Data\") to a network shared folder (e.g., "\\Server\Shared\Data\")
            // Adjust the strings below to match your actual source and target paths.
            string newPath = originalPath.Replace(
                @"C:\Data\",
                @"\\Server\Shared\Data\");

            // Update the external link with the new network path
            link.OriginalDataSource = newPath;
        }

        // Save the modified workbook (replace with your desired output path)
        string outputFile = "output.xls";
        workbook.Save(outputFile);
    }
}