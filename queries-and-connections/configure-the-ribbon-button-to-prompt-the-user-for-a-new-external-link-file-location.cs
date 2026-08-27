// Title: Create a custom Ribbon button in Aspose.Cells for .NET that prompts for an external workbook path and adds selected sheet links
// AI Prompts: Write C# code using Aspose.Cells to embed Ribbon XML with a button that requests the external workbook file path and a comma‑separated list of sheet names via console input, then adds those sheets as external links to the workbook. | Show how to save the workbook as a macro‑enabled .xlsm file so the custom Ribbon UI and the added external links persist after closing.
// Common Searches: how to embed a custom ribbon button in Aspose.Cells .NET to add external workbook links | prompt user for external Excel file path and sheet names when creating external links with Aspose.Cells | save Aspose.Cells workbook with RibbonXml as macro enabled file | Aspose.Cells C# add external links from UI input | configure RibbonXml property to launch a macro that adds external links in Aspose.Cells
// Tags: Aspose.Cells custom ribbon XML button | ExternalLinks.Add workbook path C# | macro-enabled .xlsm save Aspose.Cells | console input external link Aspose.Cells | add external sheet links Aspose.Cells .NET

using System;
using Aspose.Cells;

// The example creates a new Workbook, assigns custom Ribbon XML containing a button, reads an external workbook path and a comma‑separated list of sheet names from the console, adds those sheets as external links via Worksheets.ExternalLinks.Add, and saves the file as a macro‑enabled .xlsm so the Ribbon UI and links are retained.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Define Ribbon XML with a button that would invoke a macro (placeholder name)
        string ribbonXml =
            "<customUI xmlns=\"http://schemas.microsoft.com/office/2006/01/customui\">" +
            "  <ribbon>" +
            "    <tabs>" +
            "      <tab id=\"customTab\" label=\"External Links\">" +
            "        <group id=\"linkGroup\" label=\"Links\">" +
            "          <button id=\"addLinkBtn\" label=\"Add External Link\" size=\"large\" onAction=\"AddExternalLink\" />" +
            "        </group>" +
            "      </tab>" +
            "    </tabs>" +
            "  </ribbon>" +
            "</customUI>";

        // Set the RibbonXml property (rule usage)
        workbook.RibbonXml = ribbonXml;

        // Prompt the user for the external workbook file path
        Console.WriteLine("Enter the full path of the external workbook:");
        string externalFilePath = Console.ReadLine();

        // Prompt the user for sheet names to link (comma‑separated)
        Console.WriteLine("Enter sheet names to link (comma separated, e.g., Sheet1,Sheet2):");
        string sheetsInput = Console.ReadLine();
        string[] sheetNames = sheetsInput.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
        for (int i = 0; i < sheetNames.Length; i++)
        {
            sheetNames[i] = sheetNames[i].Trim();
        }

        // Add the external link using the ExternalLinkCollection.Add method (rule usage)
        int linkIndex = workbook.Worksheets.ExternalLinks.Add(externalFilePath, sheetNames);
        Console.WriteLine($"External link added at index {linkIndex}.");

        // Save the workbook (macro‑enabled to retain Ribbon XML)
        workbook.Save("WorkbookWithRibbon.xlsm");
    }
}
