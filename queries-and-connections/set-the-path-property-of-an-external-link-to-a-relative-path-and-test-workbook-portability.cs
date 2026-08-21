// Title: Make Excel External Links Portable with Aspose.Cells – Set Relative DataSource in C#
// Description: Demonstrates how to create an external workbook, assign a relative path to ExternalLink.DataSource, save the source file, reload it, and verify that the link resolves correctly using Aspose.Cells for .NET.
// Keywords: Aspose.Cells external link relative path | C# Excel workbook portability | ExternalLink DataSource Aspose | relative external reference .NET | verify external link after save
// Common Searches: Aspose.Cells set external link to relative path | make Excel links portable C# | how to test external link path type Aspose.Cells | combine workbook folder with relative external link | validate external workbook reference after moving file
// Developer Intent: Set ExternalLink.DataSource to a relative path and ensure the workbook can locate the external file after being saved and reloaded.
// Use Cases: Create a template that references data in a sibling folder, using a relative path for cross‑environment deployment. | Programmatically adjust existing external links to relative paths to improve file portability. | Automated testing of Excel workbooks to confirm that relative external references resolve to actual files at runtime.
// AI Prompts: Generate C# code with Aspose.Cells that changes an external link's DataSource to "..\External\ExternalData.xlsx" and checks that PathType is Relative. | Explain how to combine a workbook's directory with a relative ExternalLink.DataSource to obtain the absolute file path in .NET. | Show a step‑by‑step verification that a workbook saved with a relative external link can still find the external file after being moved to a different folder.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExternalLinkPortabilityDemo
{
    // Demonstrates how to create an external workbook, assign a relative path to ExternalLink.DataSource, save the source file, reload it, and verify that the link resolves correctly using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // Base directory for the demo
                string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "PortabilityDemo");
                string externalDir = Path.Combine(baseDir, "External");
                string sourceDir = Path.Combine(baseDir, "Source");

                // Ensure clean folders
                if (Directory.Exists(baseDir))
                    Directory.Delete(baseDir, true);
                Directory.CreateDirectory(externalDir);
                Directory.CreateDirectory(sourceDir);

                // -------------------------------------------------
                // 1. Create the external workbook (the data source)
                // -------------------------------------------------
                Workbook externalWb = new Workbook();
                Worksheet extSheet = externalWb.Worksheets[0];
                extSheet.Name = "Data";
                extSheet.Cells["A1"].PutValue(12345); // sample data
                string externalFileName = "ExternalData.xlsx";
                string externalFullPath = Path.Combine(externalDir, externalFileName);
                externalWb.Save(externalFullPath);

                // -------------------------------------------------
                // 2. Create the source workbook that will reference the external file
                // -------------------------------------------------
                Workbook sourceWb = new Workbook();
                Worksheet srcSheet = sourceWb.Worksheets[0];
                srcSheet.Name = "Main";

                // Add a formula that references the external workbook.
                srcSheet.Cells["A1"].Formula = $"='[ExternalData.xlsx]Data'!A1";

                // Ensure the external link was created
                if (sourceWb.Worksheets.ExternalLinks.Count == 0)
                    throw new InvalidOperationException("External link was not created.");

                // Retrieve the automatically created external link
                ExternalLink link = sourceWb.Worksheets.ExternalLinks[0];

                // -------------------------------------------------
                // 3. Set the external link's DataSource to a relative path
                // -------------------------------------------------
                // Relative path from the source workbook location (Source folder) to the external workbook (External folder)
                string relativePath = Path.Combine("..", "External", externalFileName);
                link.DataSource = relativePath; // now the link uses a relative path

                // Display the path type (should be Relative)
                Console.WriteLine($"Path Type after setting relative path: {link.PathType}");

                // -------------------------------------------------
                // 4. Save the source workbook
                // -------------------------------------------------
                string sourceFileName = "SourceWorkbook.xlsx";
                string sourceFullPath = Path.Combine(sourceDir, sourceFileName);
                sourceWb.Save(sourceFullPath);

                // -------------------------------------------------
                // 5. Load the saved source workbook from its location and verify portability
                // -------------------------------------------------
                if (!File.Exists(sourceFullPath))
                    throw new FileNotFoundException("Saved source workbook not found.", sourceFullPath);

                Workbook loadedSource = new Workbook(sourceFullPath);

                if (loadedSource.Worksheets.ExternalLinks.Count == 0)
                    throw new InvalidOperationException("No external links found in the loaded workbook.");

                ExternalLink loadedLink = loadedSource.Worksheets.ExternalLinks[0];
                Console.WriteLine($"Loaded External Link DataSource (relative): {loadedLink.DataSource}");

                // Combine the workbook's directory with the relative link to obtain the full path at runtime
                string sourceFolder = Path.GetDirectoryName(sourceFullPath);
                if (sourceFolder == null)
                    throw new InvalidOperationException("Unable to determine the source workbook directory.");

                string combinedFullPath = Path.GetFullPath(Path.Combine(sourceFolder, loadedLink.DataSource));
                Console.WriteLine($"Combined full path to external file: {combinedFullPath}");

                // Verify that the combined path points to an existing file
                bool externalExists = File.Exists(combinedFullPath);
                Console.WriteLine($"External file exists at combined path: {externalExists}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
