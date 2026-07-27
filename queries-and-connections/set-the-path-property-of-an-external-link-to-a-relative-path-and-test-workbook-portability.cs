using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExternalLinkPortability
{
    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Prepare a temporary folder for the test files
            // -----------------------------------------------------------------
            string baseFolder = Path.Combine(Path.GetFullPath("TestPortability"));
            Directory.CreateDirectory(baseFolder);

            // -----------------------------------------------------------------
            // 2. Create the external workbook that will be linked to
            // -----------------------------------------------------------------
            string externalFileName = "ExternalData.xlsx";
            string externalFilePath = Path.Combine(baseFolder, externalFileName);

            Workbook externalWb = new Workbook();
            Worksheet extSheet = externalWb.Worksheets[0];
            extSheet.Name = "Data";
            extSheet.Cells["A1"].PutValue(12345); // sample data
            externalWb.Save(externalFilePath);   // Save the external workbook

            // -----------------------------------------------------------------
            // 3. Create the main workbook and add a formula that references the external file
            // -----------------------------------------------------------------
            Workbook mainWb = new Workbook();
            Worksheet mainSheet = mainWb.Worksheets[0];
            mainSheet.Name = "Main";

            // Use only the file name (relative path) in the formula
            mainSheet.Cells["A1"].Formula = $"='[{externalFileName}]Data'!A1";

            // At this point Aspose.Cells automatically creates an ExternalLink entry.
            // Retrieve that link and explicitly set its DataSource to a relative path.
            ExternalLink link = mainWb.Worksheets.ExternalLinks[0];
            // Ensure the DataSource is a relative path (just the file name)
            link.DataSource = externalFileName; // relative path

            // Save the main workbook in the same folder
            string mainFileName = "MainWorkbook.xlsx";
            string mainFilePath = Path.Combine(baseFolder, mainFileName);
            mainWb.Save(mainFilePath);

            // -----------------------------------------------------------------
            // 4. Simulate moving the whole folder to a new location (portability test)
            // -----------------------------------------------------------------
            string movedFolder = Path.Combine(Path.GetFullPath("TestPortabilityMoved"));
            Directory.CreateDirectory(movedFolder);

            // Copy both files to the new location
            File.Copy(externalFilePath, Path.Combine(movedFolder, externalFileName), true);
            File.Copy(mainFilePath, Path.Combine(movedFolder, mainFileName), true);

            // -----------------------------------------------------------------
            // 5. Load the moved main workbook and verify that the external link
            //    still uses the relative path and can be resolved.
            // -----------------------------------------------------------------
            string movedMainPath = Path.Combine(movedFolder, mainFileName);
            Workbook loadedMain = new Workbook(movedMainPath);

            // The external link collection should contain one entry
            if (loadedMain.Worksheets.ExternalLinks.Count > 0)
            {
                ExternalLink loadedLink = loadedMain.Worksheets.ExternalLinks[0];
                Console.WriteLine("External link DataSource after move: " + loadedLink.DataSource);
                Console.WriteLine("Path type: " + loadedLink.PathType); // should be Relative
            }
            else
            {
                Console.WriteLine("No external links found after moving the workbook.");
            }

            // Optional: read the value from the linked cell to prove it works
            // (Aspose.Cells does not automatically evaluate external references,
            //  but we can manually open the external workbook and read the cell.)
            Workbook loadedExternal = new Workbook(Path.Combine(movedFolder, externalFileName));
            Console.WriteLine("Value in external workbook A1: " + loadedExternal.Worksheets["Data"].Cells["A1"].Value);
        }
    }
}