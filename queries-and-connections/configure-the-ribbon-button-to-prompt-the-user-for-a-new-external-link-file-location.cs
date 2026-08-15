// Title: Create a Ribbon button that prompts for a workbook path and adds an external link using Aspose.Cells (C#)
// Description: Demonstrates how to embed custom Ribbon XML in a workbook, display a console prompt for an external workbook file, add the selected sheet as an external link via Worksheets.ExternalLinks, and save the result as a macro‑enabled .xlsm file with the Ribbon retained.
// Keywords: Aspose.Cells C# Ribbon XML | custom Ribbon button Aspose.Cells | add external link workbook Aspose.Cells | prompt user for file path C# | macro enabled workbook save | Worksheets.ExternalLinks | Excel custom UI Aspose | external workbook linking
// Common Searches: Aspose.Cells add external link from Ribbon button | C# prompt for workbook path and link it with Aspose.Cells | How to embed custom Ribbon XML in an Aspose.Cells workbook | Save workbook with Ribbon as .xlsm using Aspose.Cells | Create external links collection programmatically Aspose.Cells
// Developer Intent: Add a Ribbon UI element that asks the user for a workbook location and creates an external link to that workbook in the current file.
// Use Cases: Provide end‑users with a one‑click UI to link another workbook at runtime. | Programmatically add external links after validating a user‑entered file path. | Distribute macro‑enabled Excel files that retain custom Ribbon tabs and external link functionality.
// AI Prompts: Generate C# code that uses Aspose.Cells to add a custom Ribbon button which opens a file‑picker dialog and creates an external link to the selected workbook. | Write robust error handling for Worksheets.ExternalLinks.Add when the file path is invalid or the specified sheet does not exist. | Show how to embed custom Ribbon XML in a workbook and preserve it after saving as an .xlsm file with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to embed custom Ribbon XML in a workbook, display a console prompt for an external workbook file, add the selected sheet as an external link via Worksheets.ExternalLinks, and save the result as a macro‑enabled .xlsm file with the Ribbon retained.
class ConfigureRibbonAndAddExternalLink
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Define custom Ribbon XML with a button that will be used to add an external link.
        // In a real Excel UI this button could be linked to a macro; here we simulate the action in code.
        string ribbonXml =
            "<customUI xmlns=\"http://schemas.microsoft.com/office/2006/01/customui\">" +
            "  <ribbon>" +
            "    <tabs>" +
            "      <tab id=\"customTab\" label=\"External Links\">" +
            "        <group id=\"linkGroup\" label=\"Link Operations\">" +
            "          <button id=\"addLinkButton\" label=\"Add External Link\" size=\"large\" />" +
            "        </group>" +
            "      </tab>" +
            "    </tabs>" +
            "  </ribbon>" +
            "</customUI>";

        // Assign the Ribbon XML to the workbook
        workbook.RibbonXml = ribbonXml;

        // Prompt the user for the external workbook file location
        Console.WriteLine("Enter the full path of the external workbook to link:");
        string externalFilePath = Console.ReadLine();

        // Validate input (basic check)
        if (string.IsNullOrWhiteSpace(externalFilePath))
        {
            Console.WriteLine("No file path provided. Exiting.");
            return;
        }

        // Define which sheets (or ranges) from the external workbook should be linked.
        // For demonstration we link the entire first sheet.
        string[] sheetNames = new string[] { "Sheet1" };

        try
        {
            // Add the external link to the workbook's external links collection
            int linkIndex = workbook.Worksheets.ExternalLinks.Add(externalFilePath, sheetNames);
            Console.WriteLine($"External link added at index {linkIndex}.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error adding external link: {ex.Message}");
        }

        // Save the workbook as a macro‑enabled file so the custom Ribbon is retained
        workbook.Save("WorkbookWithRibbonAndLink.xlsm");
        Console.WriteLine("Workbook saved as 'WorkbookWithRibbonAndLink.xlsm'.");
    }
}
