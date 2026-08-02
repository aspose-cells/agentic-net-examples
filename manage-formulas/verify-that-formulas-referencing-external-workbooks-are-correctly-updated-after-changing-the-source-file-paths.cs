// Title: Update external workbook links after moving files – Aspose.Cells for .NET example
// Description: Demonstrates how to create an external workbook, reference it from a main workbook, move the source file, change the ExternalLink.DataSource, refresh linked data with UpdateLinkedDataSource, and recalculate formulas to verify the new value.
// Keywords: Aspose.Cells external link update | change external workbook path .NET | Refresh linked formulas Aspose.Cells | UpdateLinkedDataSource example | ExternalLink.DataSource C#
// Common Searches: how to change external link path in Aspose.Cells | refresh formulas after moving Excel file Aspose | update external workbook reference .NET | Aspose.Cells recalculate after external file rename | C# example for updating external workbook links
// Developer Intent: Confirm that formulas referencing an external workbook continue to return correct values after the source file location is changed.
// Use Cases: Modify ExternalLink.DataSource to point to a new file location. | Call UpdateLinkedDataSource to reload data from the moved workbook. | Recalculate formulas to ensure linked cells reflect the updated source.
// AI Prompts: Show C# code that changes the data source of an external link in Aspose.Cells and refreshes linked formulas. | Provide a step‑by‑step example verifying that a formula referencing an external workbook returns the correct value after the source file is moved. | Explain how UpdateLinkedDataSource works with multiple external workbooks in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to create an external workbook, reference it from a main workbook, move the source file, change the ExternalLink.DataSource, refresh linked data with UpdateLinkedDataSource, and recalculate formulas to verify the new value.
class VerifyExternalLinkUpdate
{
    static void Main()
    {
        // ------------------------------------------------------------
        // 1. Create an external workbook and write a value to A1
        // ------------------------------------------------------------
        Workbook externalWb = new Workbook();
        externalWb.Worksheets[0].Cells["A1"].PutValue("Original Value");
        string originalPath = Path.Combine(Directory.GetCurrentDirectory(), "ExternalOriginal.xlsx");
        externalWb.Save(originalPath); // save rule

        // ------------------------------------------------------------
        // 2. Create the main workbook that references the external file
        // ------------------------------------------------------------
        Workbook mainWb = new Workbook();
        Worksheet mainSheet = mainWb.Worksheets[0];

        // Formula that points to the external workbook (file name only)
        mainSheet.Cells["A1"].Formula = $"='[{Path.GetFileName(originalPath)}]Sheet1'!A1";

        // Register the external link so Aspose knows where to look for the source file
        mainWb.Worksheets.ExternalLinks.Add(originalPath, new string[] { "Sheet1" });

        // Calculate to obtain the initial value from the original external workbook
        mainWb.CalculateFormula();
        Console.WriteLine("Initial value: " + mainSheet.Cells["A1"].StringValue);

        // ------------------------------------------------------------
        // 3. Simulate moving the external workbook to a new location
        // ------------------------------------------------------------
        string newPath = Path.Combine(Directory.GetCurrentDirectory(), "ExternalNew.xlsx");
        externalWb.Save(newPath); // save the same content under a new name

        // Update the external link's data source to point to the new file location
        ExternalLink extLink = mainWb.Worksheets.ExternalLinks[0];
        extLink.DataSource = newPath; // change path

        // Load the new external workbook instance
        Workbook newExternalWb = new Workbook(newPath);

        // Tell the main workbook to refresh its external data from the new source
        mainWb.UpdateLinkedDataSource(new Workbook[] { newExternalWb });

        // Recalculate formulas after the link update
        mainWb.CalculateFormula();
        Console.WriteLine("Updated value after path change: " + mainSheet.Cells["A1"].StringValue);

        // ------------------------------------------------------------
        // 4. Save the main workbook (demonstrates the save rule)
        // ------------------------------------------------------------
        mainWb.Save("MainWorkbookResult.xlsx");
    }
}
