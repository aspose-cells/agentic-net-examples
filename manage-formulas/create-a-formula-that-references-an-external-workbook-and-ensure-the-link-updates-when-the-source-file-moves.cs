using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExternalLinkDemo
{
    class Program
    {
        static void Main()
        {
            // Paths for the original and moved external workbook
            string baseDir = Path.Combine(Path.GetTempPath(), "AsposeExternalLinkDemo");
            string originalDir = Path.Combine(baseDir, "Original");
            string movedDir = Path.Combine(baseDir, "Moved");
            Directory.CreateDirectory(originalDir);
            Directory.CreateDirectory(movedDir);

            // File names
            string externalFileName = "ExternalWorkbook.xlsx";
            string externalOriginalPath = Path.Combine(originalDir, externalFileName);
            string externalMovedPath = Path.Combine(movedDir, externalFileName);
            string mainFilePath = Path.Combine(baseDir, "MainWorkbook.xlsx");

            // -------------------------------------------------
            // 1. Create the external workbook and set a value
            // -------------------------------------------------
            Workbook externalWb = new Workbook();
            Worksheet extSheet = externalWb.Worksheets[0];
            extSheet.Name = "Sheet1";
            extSheet.Cells["A1"].PutValue("Original Value");
            externalWb.Save(externalOriginalPath, SaveFormat.Xlsx);
            externalWb.Dispose();

            // -------------------------------------------------
            // 2. Create the main workbook and add a formula that references the external workbook
            // -------------------------------------------------
            Workbook mainWb = new Workbook();
            Worksheet mainSheet = mainWb.Worksheets[0];
            // Formula uses only the file name; Aspose.Cells will resolve it relative to the workbook location
            mainSheet.Cells["A1"].Formula = $"=[{externalFileName}]Sheet1!A1";
            mainWb.Save(mainFilePath, SaveFormat.Xlsx);

            // -------------------------------------------------
            // 3. Simulate moving the external workbook to a new folder
            // -------------------------------------------------
            File.Copy(externalOriginalPath, externalMovedPath, true);
            // Update the external link's data source to the new location
            // The link is stored in the main workbook's ExternalLinks collection
            ExternalLink link = mainWb.Worksheets.ExternalLinks[0];
            link.DataSource = externalMovedPath; // set full path to the moved file

            // -------------------------------------------------
            // 4. Load the moved external workbook and refresh the link
            // -------------------------------------------------
            Workbook movedExternalWb = new Workbook(externalMovedPath);
            // Optionally change the value to demonstrate that the link updates
            movedExternalWb.Worksheets[0].Cells["A1"].PutValue("Updated Value");
            movedExternalWb.Save(externalMovedPath, SaveFormat.Xlsx);

            // Update the main workbook with the new external workbook instance
            mainWb.UpdateLinkedDataSource(new Workbook[] { movedExternalWb });

            // Recalculate formulas so the new value is reflected
            mainWb.CalculateFormula();

            // -------------------------------------------------
            // 5. Output the result
            // -------------------------------------------------
            Console.WriteLine("Value in Main Workbook A1 after link update: " +
                              mainSheet.Cells["A1"].StringValue);

            // Clean up (optional)
            mainWb.Dispose();
            movedExternalWb.Dispose();
        }
    }
}