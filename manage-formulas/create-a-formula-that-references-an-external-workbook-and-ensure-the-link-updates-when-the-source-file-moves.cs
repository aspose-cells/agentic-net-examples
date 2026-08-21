// Title: Update an external workbook link in Aspose.Cells .NET after moving the source file
// Description: Demonstrates how to create an external workbook, reference it in a main workbook using a formula, add the link to the ExternalLinks collection, move the source file, change the link's DataSource to the new path, refresh linked data with UpdateLinkedDataSource, recalculate formulas, and save the updated workbook.
// Keywords: Aspose.Cells external link | C# external workbook reference | Update DataSource Aspose.Cells | Refresh linked data source | CalculateFormula after link change | ExternalLinks collection .NET | move source workbook Aspose
// Common Searches: how to change external link path in Aspose.Cells | update formula reference after moving workbook .NET | refresh external data source Aspose.Cells C# | set external link without full path Aspose.Cells | Aspose.Cells update external workbook location
// Developer Intent: Programmatically modify an external workbook reference so the formula continues to work after the source file is relocated.
// Use Cases: Link a master workbook to a value in another workbook and keep the link valid when the source file is moved. | Change the DataSource of an ExternalLink at runtime after a file system reorganization. | Refresh linked data and recalculate dependent formulas to reflect the latest values from a moved workbook.
// AI Prompts: Generate C# code with Aspose.Cells that adds an external link to a workbook and updates its DataSource after moving the source file. | Show how to refresh linked data sources and recalculate formulas in Aspose.Cells when the external workbook path changes. | Explain the role of the ExternalLinks collection and DataSource property in Aspose.Cells and how to programmatically update them.

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to create an external workbook, reference it in a main workbook using a formula, add the link to the ExternalLinks collection, move the source file, change the link's DataSource to the new path, refresh linked data with UpdateLinkedDataSource, recalculate formulas, and save the updated workbook.
class ExternalLinkUpdateDemo
{
    static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create an external workbook and save it to an initial location.
        // -----------------------------------------------------------------
        string originalPath = Path.Combine(Directory.GetCurrentDirectory(), "ExternalOriginal.xlsx");
        Workbook externalWb = new Workbook();
        externalWb.Worksheets[0].Cells["A1"].PutValue("Original Value");
        externalWb.Save(originalPath);

        // ---------------------------------------------------------------
        // 2. Create the main workbook and set a formula that references the external file.
        // ---------------------------------------------------------------
        Workbook mainWb = new Workbook();
        Worksheet mainSheet = mainWb.Worksheets[0];
        // Formula uses only the file name; the actual path is stored in the ExternalLinks collection.
        mainSheet.Cells["A1"].Formula = $"='[{Path.GetFileName(originalPath)}]Sheet1'!A1";

        // Add an entry to the ExternalLinks collection so Aspose knows about the external source.
        int linkIndex = mainWb.Worksheets.ExternalLinks.Add(originalPath, new string[] { "Sheet1" });

        // ---------------------------------------------------------------
        // 3. Simulate moving the external workbook to a new folder.
        // ---------------------------------------------------------------
        string newFolder = Path.Combine(Directory.GetCurrentDirectory(), "MovedFolder");
        Directory.CreateDirectory(newFolder);
        string newPath = Path.Combine(newFolder, "ExternalMoved.xlsx");
        File.Copy(originalPath, newPath, true);
        File.Delete(originalPath); // optional: delete the original to emulate a move

        // ---------------------------------------------------------------
        // 4. Update the external link's DataSource to point to the new location.
        // ---------------------------------------------------------------
        ExternalLink link = mainWb.Worksheets.ExternalLinks[linkIndex];
        link.DataSource = newPath;

        // ---------------------------------------------------------------
        // 5. Load the moved external workbook and refresh the link.
        // ---------------------------------------------------------------
        Workbook movedExternal = new Workbook(newPath);
        mainWb.UpdateLinkedDataSource(new Workbook[] { movedExternal });

        // ---------------------------------------------------------------
        // 6. Recalculate formulas so the updated value is reflected.
        // ---------------------------------------------------------------
        mainWb.CalculateFormula();

        // ---------------------------------------------------------------
        // 7. Display the result and save the main workbook.
        // ---------------------------------------------------------------
        Console.WriteLine("Updated value in main workbook: " + mainSheet.Cells["A1"].StringValue);
        mainWb.Save("MainWithUpdatedLink.xlsx");
    }
}
