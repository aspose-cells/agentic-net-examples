using System;
using System.IO;
using Aspose.Cells;

namespace ExternalLinkRelativePathDemo
{
    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Prepare an external workbook in a subfolder called "Data"
            // -----------------------------------------------------------------
            string baseDir = Directory.GetCurrentDirectory();
            string dataDir = Path.Combine(baseDir, "Data");
            Directory.CreateDirectory(dataDir);

            string externalFileName = "ExternalData.xlsx";
            string externalFilePath = Path.Combine(dataDir, externalFileName);

            // Create the external workbook and add a value
            Workbook externalWb = new Workbook();
            externalWb.Worksheets[0].Cells["A1"].PutValue(12345);
            externalWb.Save(externalFilePath);

            // -----------------------------------------------------------------
            // 2. Create the main workbook and add a formula that references the external file using a relative path
            // -----------------------------------------------------------------
            Workbook mainWb = new Workbook();

            // Relative path from the main workbook location to the external file
            string relativePath = Path.Combine("Data", externalFileName).Replace('\\', '/');

            // Set formula that points to the external workbook
            // Note: The formula syntax for external references is:
            // ='[relativePath]SheetName'!CellReference
            mainWb.Worksheets[0].Cells["A1"].Formula = $"='[{relativePath}]Sheet1'!A1";

            // At this point Aspose.Cells automatically creates an ExternalLink object.
            // Ensure the DataSource of the external link is the relative path we expect.
            if (mainWb.Worksheets.ExternalLinks.Count > 0)
            {
                ExternalLink link = mainWb.Worksheets.ExternalLinks[0];
                // Override the DataSource with the relative path (if not already set)
                link.DataSource = relativePath;
                Console.WriteLine("External link DataSource set to relative path: " + link.DataSource);
                Console.WriteLine("Path Type: " + link.PathType);
            }

            // Save the main workbook in the base directory
            string mainFileName = "MainWorkbook.xlsx";
            string mainFilePath = Path.Combine(baseDir, mainFileName);
            mainWb.Save(mainFilePath);
            Console.WriteLine("Main workbook saved to: " + mainFilePath);

            // -----------------------------------------------------------------
            // 3. Test portability: copy both files to a new temporary folder and load the main workbook
            // -----------------------------------------------------------------
            string portableDir = Path.Combine(baseDir, "PortableTest");
            Directory.CreateDirectory(portableDir);

            // Copy external workbook preserving relative folder structure
            string portableDataDir = Path.Combine(portableDir, "Data");
            Directory.CreateDirectory(portableDataDir);
            File.Copy(externalFilePath, Path.Combine(portableDataDir, externalFileName), true);

            // Copy main workbook
            string portableMainPath = Path.Combine(portableDir, mainFileName);
            File.Copy(mainFilePath, portableMainPath, true);

            // Load the workbook from the new location
            Workbook loadedWb = new Workbook(portableMainPath);

            // Verify that the external link still points to the relative path
            if (loadedWb.Worksheets.ExternalLinks.Count > 0)
            {
                ExternalLink loadedLink = loadedWb.Worksheets.ExternalLinks[0];
                Console.WriteLine("After moving, external link DataSource: " + loadedLink.DataSource);
                Console.WriteLine("Path Type after moving: " + loadedLink.PathType);
            }

            // Optionally, read the value from the external link to ensure it resolves correctly
            // (Aspose.Cells does not automatically evaluate external references without loading the external file,
            //  but we can demonstrate that the path is correct.)
            Console.WriteLine("Portability test completed.");
        }
    }
}