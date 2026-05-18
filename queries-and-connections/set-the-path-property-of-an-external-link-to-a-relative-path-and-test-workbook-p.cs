using System;
using System.IO;
using Aspose.Cells;

class ExternalLinkRelativePathDemo
{
    static void Main()
    {
        // Create temporary folders for the demo
        string baseDir = Path.Combine(Path.GetTempPath(), "AsposeDemo");
        string sourceDir = Path.Combine(baseDir, "Source");
        string externalDir = Path.Combine(baseDir, "External");
        Directory.CreateDirectory(sourceDir);
        Directory.CreateDirectory(externalDir);

        // -----------------------------------------------------------------
        // 1. Create an external workbook that will be referenced later
        // -----------------------------------------------------------------
        Workbook externalWb = new Workbook();
        externalWb.Worksheets[0].Cells["A1"].PutValue(123); // sample data
        string externalFilePath = Path.Combine(externalDir, "ExternalData.xlsx");
        externalWb.Save(externalFilePath); // save the external workbook

        // -----------------------------------------------------------------
        // 2. Create the main workbook and add a relative external link
        // -----------------------------------------------------------------
        Workbook sourceWb = new Workbook();

        // Build a relative path that goes up one directory and then into the External folder
        // Example relative path: "..\External\ExternalData.xlsx"
        string relativePath = Path.Combine("..", "External", "ExternalData.xlsx");

        // Add the external link using DirectoryType.UpDirectory to indicate a relative up‑directory reference
        int linkIndex = sourceWb.Worksheets.ExternalLinks.Add(
            DirectoryType.UpDirectory,   // relative up‑directory
            relativePath,                // the relative file name
            new string[] { "Sheet1" }    // sheets to reference
        );

        // Use the external link in a formula (the file name part is the file name only)
        sourceWb.Worksheets[0].Cells["A1"].Formula = $"='[{Path.GetFileName(relativePath)}]Sheet1'!A1";

        // Save the main workbook
        string sourceFilePath = Path.Combine(sourceDir, "SourceWorkbook.xlsx");
        sourceWb.Save(sourceFilePath);

        // -----------------------------------------------------------------
        // 3. Load the saved workbook and verify that the link is stored as relative
        // -----------------------------------------------------------------
        Workbook loadedWb = new Workbook(sourceFilePath);
        ExternalLink link = loadedWb.Worksheets.ExternalLinks[0];

        Console.WriteLine("External link DataSource (should be relative): " + link.DataSource);
        Console.WriteLine("External link PathType: " + link.PathType);

        // -----------------------------------------------------------------
        // (Optional) Clean up temporary files/folders
        // -----------------------------------------------------------------
        // Directory.Delete(baseDir, true);
    }
}