using System;
using Aspose.Cells;

namespace UpdateExternalReference
{
    class Program
    {
        static void Main()
        {
            // Load the workbook that contains the external reference
            Workbook workbook = new Workbook("OriginalWorkbook.xlsx");

            // Define old and new file paths (adjust as needed)
            string oldPath = @"C:\OldFolder\ExternalData.xlsx";
            string newPath = @"D:\NewFolder\ExternalData.xlsx";

            // ----- Update ExternalLink objects -----
            // Each external link stores its data source (file path). Replace the old path with the new one.
            ExternalLinkCollection externalLinks = workbook.Worksheets.ExternalLinks;
            for (int i = 0; i < externalLinks.Count; i++)
            {
                ExternalLink link = externalLinks[i];
                if (!string.IsNullOrEmpty(link.DataSource) && link.DataSource.Contains(oldPath))
                {
                    // Update the data source to the new location
                    link.DataSource = link.DataSource.Replace(oldPath, newPath);
                }
            }

            // ----- Update Named Ranges that refer to the external file -----
            // Named ranges store their reference as a formula string (e.g., "='[ExternalData.xlsx]Sheet1'!$A$1").
            // Replace occurrences of the old path within those formulas.
            foreach (Name namedRange in workbook.Worksheets.Names)
            {
                if (!string.IsNullOrEmpty(namedRange.RefersTo) && namedRange.RefersTo.Contains(oldPath))
                {
                    string updatedRefersTo = namedRange.RefersTo.Replace(oldPath, newPath);
                    namedRange.RefersTo = updatedRefersTo;
                }
            }

            // Save the modified workbook
            workbook.Save("UpdatedWorkbook.xlsx");
        }
    }
}