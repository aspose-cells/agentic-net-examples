// Title: Update and Refresh External Workbook Links After Moving the Source File – Aspose.Cells for .NET Example
// Description: Demonstrates how to create an external workbook, link to it from a main workbook, move the source file to a new folder, update the ExternalLink.DataSource and OriginalDataSource paths, refresh the linked data with UpdateLinkedDataSource, recalculate formulas, and verify that the main workbook reflects the new value.
// Keywords: Aspose.Cells | C# | .NET | external workbook link | ExternalLink DataSource | OriginalDataSource | UpdateLinkedDataSource | recalculate formulas | move external file | refresh linked data | formula update after file move
// Common Searches: Aspose.Cells update external link after moving file | Refresh formulas that reference external workbook C# | Change ExternalLink DataSource path Aspose.Cells | Update linked data source for moved workbook .NET | Recalculate workbook formulas after external file relocation
// Developer Intent: Ensure that formulas referencing an external workbook continue to work after the source workbook has been moved to a different location.
// Use Cases: Create an external workbook, write a value, and save it. | Insert a formula in a main workbook that points to the external file. | Move the external workbook to another directory and update the link's DataSource and OriginalDataSource. | Call UpdateLinkedDataSource with the moved workbook, then recalculate formulas to reflect the new data. | Read the cell value to confirm the update and optionally save the main workbook.
// AI Prompts: Generate C# code using Aspose.Cells that updates the DataSource of an ExternalLink after the external workbook is moved. | Show how to refresh formulas dependent on external workbooks when their file paths change in Aspose.Cells for .NET. | Explain the steps to verify that a main workbook reflects the updated value from a moved external workbook using UpdateLinkedDataSource and CalculateFormula.

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to create an external workbook, link to it from a main workbook, move the source file to a new folder, update the ExternalLink.DataSource and OriginalDataSource paths, refresh the linked data with UpdateLinkedDataSource, recalculate formulas, and verify that the main workbook reflects the new value.
class VerifyExternalLinkUpdate
{
    static void Main()
    {
        // ------------------------------------------------------------
        // 1. Create an external workbook and write a value to A1.
        // ------------------------------------------------------------
        Workbook externalWb = new Workbook();
        externalWb.Worksheets[0].Cells["A1"].PutValue("Original Value");
        string originalExternalPath = "ExternalWorkbook.xlsx";
        externalWb.Save(originalExternalPath);

        // ------------------------------------------------------------
        // 2. Create the main workbook that references the external file.
        // ------------------------------------------------------------
        Workbook mainWb = new Workbook();
        Worksheet mainSheet = mainWb.Worksheets[0];
        // Formula points to the external workbook we just created.
        mainSheet.Cells["A1"].Formula = $"=[{originalExternalPath}]Sheet1!A1";

        // Verify that an external link entry was automatically added.
        Console.WriteLine("External links count before path change: " +
                          mainWb.Worksheets.ExternalLinks.Count);

        // ------------------------------------------------------------
        // 3. Simulate moving the external workbook to a new folder.
        //    Update the stored link path accordingly.
        // ------------------------------------------------------------
        string newExternalPath = Path.Combine("MovedFolder", "ExternalWorkbook.xlsx");
        // Ensure the target folder exists.
        Directory.CreateDirectory("MovedFolder");
        // Copy the original file to the new location.
        File.Copy(originalExternalPath, newExternalPath, true);

        // Update the external link's DataSource (and OriginalDataSource) to the new path.
        if (mainWb.Worksheets.ExternalLinks.Count > 0)
        {
            ExternalLink link = mainWb.Worksheets.ExternalLinks[0];
            link.DataSource = newExternalPath;
            link.OriginalDataSource = newExternalPath;
        }

        // ------------------------------------------------------------
        // 4. Change the value inside the moved external workbook.
        // ------------------------------------------------------------
        Workbook movedExternal = new Workbook(newExternalPath);
        movedExternal.Worksheets[0].Cells["A1"].PutValue("Updated Value");
        movedExternal.Save(newExternalPath);

        // ------------------------------------------------------------
        // 5. Refresh the main workbook so it picks up the new external data.
        // ------------------------------------------------------------
        // Pass the workbook that contains the updated data.
        mainWb.UpdateLinkedDataSource(new Workbook[] { movedExternal });

        // Recalculate formulas to reflect the refreshed data.
        mainWb.CalculateFormula();

        // ------------------------------------------------------------
        // 6. Output the result to verify the update succeeded.
        // ------------------------------------------------------------
        Console.WriteLine("Value in main workbook after update: " + mainSheet.Cells["A1"].Value);

        // (Optional) Save the main workbook for inspection.
        mainWb.Save("MainWorkbook_WithUpdatedLink.xlsx");
    }
}
