using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExternalLinkPathUpdate
{
    class Program
    {
        static void Main()
        {
            // Define input and output paths (relative to the executable directory)
            string inputPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "input.xlsx");
            string outputPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "output.xlsx");

            // Ensure the input workbook exists; create a simple workbook if it does not.
            if (!File.Exists(inputPath))
            {
                var wb = new Workbook();
                wb.Worksheets[0].Name = "Sheet1";
                wb.Save(inputPath);
            }

            // Load the workbook that may contain external links
            Workbook workbook = new Workbook(inputPath);

            // Define the part of the old path to replace and the new absolute path segment
            string oldPathSegment = @"C:\OldFolder\";
            string newPathSegment = @"D:\NewFolder\";

            // Iterate through all external links and update their stored data source paths
            for (int i = 0; i < workbook.Worksheets.ExternalLinks.Count; i++)
            {
                ExternalLink link = workbook.Worksheets.ExternalLinks[i];

                // Update OriginalDataSource if it contains the old segment
                string original = link.OriginalDataSource;
                if (!string.IsNullOrEmpty(original) && original.Contains(oldPathSegment))
                {
                    link.OriginalDataSource = original.Replace(oldPathSegment, newPathSegment);
                }

                // Update DataSource for consistency
                string dataSource = link.DataSource;
                if (!string.IsNullOrEmpty(dataSource) && dataSource.Contains(oldPathSegment))
                {
                    link.DataSource = dataSource.Replace(oldPathSegment, newPathSegment);
                }
            }

            // Save the workbook with the modified external link paths
            workbook.Save(outputPath);
        }
    }
}