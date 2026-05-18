using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    class UpdateExternalLinksToRelativePaths
    {
        static void Main()
        {
            // Define workbook path (using current directory for demo)
            string inputPath = Path.Combine(Environment.CurrentDirectory, "MyWorkbook.xlsx");

            // Ensure the workbook exists; create a simple one if it doesn't
            if (!File.Exists(inputPath))
            {
                Workbook tempWb = new Workbook();
                tempWb.Worksheets[0].Name = "Sheet1";
                tempWb.Save(inputPath);
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Base directory for relative paths
            string workbookDirectory = Path.GetDirectoryName(inputPath);

            // Iterate through external links and convert absolute paths to relative
            ExternalLinkCollection externalLinks = workbook.Worksheets.ExternalLinks;
            for (int i = 0; i < externalLinks.Count; i++)
            {
                ExternalLink link = externalLinks[i];

                if (!string.IsNullOrEmpty(link.DataSource) && Path.IsPathRooted(link.DataSource))
                {
                    string relativePath = Path.GetRelativePath(workbookDirectory, link.DataSource);
                    link.DataSource = relativePath;
                    link.OriginalDataSource = relativePath;
                }
            }

            // Save the modified workbook
            string outputPath = Path.Combine(Environment.CurrentDirectory, "MyWorkbook_Relative.xlsx");
            workbook.Save(outputPath);
        }
    }
}