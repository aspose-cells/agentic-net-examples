// Title: Update External Workbook Links in Aspose.Cells .NET When the Source File Moves
// Description: Demonstrates how to create a main workbook that references an external workbook via a formula, register the link with the ExternalLinks collection, move the source file, change the ExternalLink.DataSource to the new path, refresh the linked data source, recalculate formulas, and save the updated workbook using Aspose.Cells for .NET.
// Keywords: Aspose.Cells external link | update external workbook reference | C# Aspose.Cells formula external file | ExternalLink.DataSource | RefreshLinkedDataSource | move external workbook | relative path formula Aspose | .NET spreadsheet linking
// Common Searches: how to change external link path in Aspose.Cells | Aspose.Cells update formula after moving source workbook | C# update ExternalLink.DataSource Aspose | refresh linked data source Aspose.Cells .NET | reference external workbook with relative path Aspose
// Developer Intent: Modify a formula’s external workbook reference so it stays valid after the source file is relocated.
// Use Cases: Create a main workbook that pulls data from another workbook using a formula and add the link to the ExternalLinks collection. | Programmatically adjust ExternalLink.DataSource when the external workbook is moved to a different folder. | Recalculate the main workbook after updating the link to reflect the new external data and save the result.
// AI Prompts: Show C# code to change the DataSource of an ExternalLink in Aspose.Cells after moving the source file. | Provide an Aspose.Cells .NET example that updates multiple external links when their files are relocated. | Explain how to ensure formulas that reference external workbooks recalculate correctly after the link path is updated.

using System;
using System.IO;
using Aspose.Cells;

namespace ExternalLinkUpdateDemo
{
    // Demonstrates how to create a main workbook that references an external workbook via a formula, register the link with the ExternalLinks collection, move the source file, change the ExternalLink.DataSource to the new path, refresh the linked data source, recalculate formulas, and save the updated workbook using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Paths for the original and moved external workbook
            string originalFolder = Path.Combine(Directory.GetCurrentDirectory(), "Original");
            string movedFolder = Path.Combine(Directory.GetCurrentDirectory(), "Moved");
            Directory.CreateDirectory(originalFolder);
            Directory.CreateDirectory(movedFolder);

            string externalFileName = "ExternalData.xlsx";
            string originalExternalPath = Path.Combine(originalFolder, externalFileName);
            string movedExternalPath = Path.Combine(movedFolder, externalFileName);

            // -----------------------------------------------------------------
            // 1. Create the external workbook and save it to the original location
            // -----------------------------------------------------------------
            Workbook externalWb = new Workbook();
            Worksheet extSheet = externalWb.Worksheets[0];
            extSheet.Name = "Sheet1";
            extSheet.Cells["A1"].PutValue("Initial Value");
            externalWb.Save(originalExternalPath, SaveFormat.Xlsx);

            // -----------------------------------------------------------------
            // 2. Create the main workbook that references the external workbook
            // -----------------------------------------------------------------
            Workbook mainWb = new Workbook();
            Worksheet mainSheet = mainWb.Worksheets[0];
            // Formula referencing the external workbook by file name (relative path)
            mainSheet.Cells["A1"].Formula = $"='[{externalFileName}]Sheet1'!A1";

            // Add the external link to the workbook's external links collection
            // This registers the link so that UpdateLinkedDataSource can work later
            mainWb.Worksheets.ExternalLinks.Add(externalFileName, new string[] { "Sheet1" });

            // Save the main workbook
            string mainPath = Path.Combine(Directory.GetCurrentDirectory(), "MainWorkbook.xlsx");
            mainWb.Save(mainPath, SaveFormat.Xlsx);

            // -----------------------------------------------------------------
            // 3. Load the main workbook and update the link using the original external workbook
            // -----------------------------------------------------------------
            Workbook loadedMain = new Workbook(mainPath);
            // Load the external workbook that the link points to
            Workbook linkedExternal = new Workbook(originalExternalPath);
            // Update the external data source so the formula can retrieve the latest value
            loadedMain.UpdateLinkedDataSource(new Workbook[] { linkedExternal });
            loadedMain.CalculateFormula();

            Console.WriteLine("Value after first calculation (original location): " +
                loadedMain.Worksheets[0].Cells["A1"].StringValue);

            // -----------------------------------------------------------------
            // 4. Simulate moving the external workbook to a new folder
            // -----------------------------------------------------------------
            File.Copy(originalExternalPath, movedExternalPath, true);
            File.Delete(originalExternalPath); // original file no longer exists

            // Change the value in the moved external workbook to demonstrate update
            Workbook movedExternal = new Workbook(movedExternalPath);
            movedExternal.Worksheets[0].Cells["A1"].PutValue("Value After Move");
            movedExternal.Save(movedExternalPath, SaveFormat.Xlsx);

            // -----------------------------------------------------------------
            // 5. Update the external link's DataSource to the new location
            // -----------------------------------------------------------------
            // The DataSource property holds the path used in the formula.
            // We replace the old path with the new absolute path.
            ExternalLink link = loadedMain.Worksheets.ExternalLinks[0];
            // Preserve the original file name part and prepend the new folder path
            link.DataSource = movedExternalPath;

            // -----------------------------------------------------------------
            // 6. Refresh the link using the moved external workbook and recalculate
            // -----------------------------------------------------------------
            Workbook newLinkedExternal = new Workbook(movedExternalPath);
            loadedMain.UpdateLinkedDataSource(new Workbook[] { newLinkedExternal });
            loadedMain.CalculateFormula();

            Console.WriteLine("Value after moving external workbook and updating link: " +
                loadedMain.Worksheets[0].Cells["A1"].StringValue);

            // -----------------------------------------------------------------
            // 7. Save the final state of the main workbook (optional)
            // -----------------------------------------------------------------
            string finalMainPath = Path.Combine(Directory.GetCurrentDirectory(), "MainWorkbook_Updated.xlsx");
            loadedMain.Save(finalMainPath, SaveFormat.Xlsx);
        }
    }
}
