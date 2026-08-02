// Title: C# – Update Excel External Link Paths to a Network Share with Aspose.Cells
// Description: This example shows how to load an Excel workbook, enumerate its ExternalLinkCollection, replace the old local folder in each link's OriginalDataSource and DataSource with a new UNC network‑share path, and save the workbook. The code follows Aspose.Cells best practices for loading and saving workbooks on Windows environments.
// Keywords: Aspose.Cells | C# | ExternalLinkCollection | update external link paths | UNC network share | Excel external references | Workbook.Save | file path migration | Excel formula links | Windows file paths
// Common Searches: how to change external link file path in Excel using Aspose.Cells C# | replace local folder with UNC path for Excel external links | update ExternalLinkCollection after moving source workbooks | Aspose.Cells example for updating external data sources | C# code to modify Excel external references to network share
// Developer Intent: Modify every external link in an Excel workbook so that formulas point to a new network‑share location.
// Use Cases: Migrate a single workbook’s external references from a local directory to a shared UNC folder. | Batch‑process multiple workbooks, applying the same path‑replacement logic to each file. | Validate that updated ExternalLink.OriginalDataSource and ExternalLink.DataSource values are correctly saved. | Integrate the path‑update routine into an automated deployment pipeline that moves source data to a central server.
// AI Prompts: Generate C# code using Aspose.Cells to replace a specific local folder in all external link paths of an Excel workbook with a UNC network share. | Explain the steps to safely update both OriginalDataSource and DataSource properties of ExternalLink objects to keep formulas consistent after moving source files. | Create a reusable method that accepts a workbook path and a new network root, updates external links, saves the workbook, and returns the output file path.

using System;
using Aspose.Cells;

namespace UpdateExternalLinksDemo
{
    // This example shows how to load an Excel workbook, enumerate its ExternalLinkCollection, replace the old local folder in each link's OriginalDataSource and DataSource with a new UNC network‑share path, and save the workbook. The code follows Aspose.Cells best practices for loading and saving workbooks on Windows environments.
    class Program
    {
        static void Main()
        {
            // Path to the workbook that contains external links
            string workbookPath = @"C:\OldFolder\MainWorkbook.xlsx";

            // New network share location where the source workbooks have been moved
            string newNetworkRoot = @"\\NetworkShare\NewFolder\";

            // Load the workbook (lifecycle rule: use load)
            Workbook workbook = new Workbook(workbookPath);

            // Iterate through all external links in the workbook
            ExternalLinkCollection externalLinks = workbook.Worksheets.ExternalLinks;
            for (int i = 0; i < externalLinks.Count; i++)
            {
                ExternalLink link = externalLinks[i];

                // OriginalDataSource holds the stored path of the external link
                string originalPath = link.OriginalDataSource;

                // Replace the old local folder with the new network share path
                // Ensure the path ends with a backslash for correct replacement
                string updatedPath = originalPath.Replace(
                    @"C:\OldFolder\", 
                    newNetworkRoot, 
                    StringComparison.OrdinalIgnoreCase);

                // Apply the updated path back to the link
                link.OriginalDataSource = updatedPath;

                // Also update DataSource to keep the formula references consistent
                link.DataSource = updatedPath;
            }

            // Save the modified workbook (lifecycle rule: use save)
            string outputPath = @"C:\OldFolder\MainWorkbook_Updated.xlsx";
            workbook.Save(outputPath);

            Console.WriteLine("External links have been updated and workbook saved to:");
            Console.WriteLine(outputPath);
        }
    }
}
