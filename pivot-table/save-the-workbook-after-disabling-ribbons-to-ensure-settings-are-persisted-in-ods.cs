// Title: Disable Ribbon UI and Save Workbook as ODS with Aspose.Cells (C#)
// Description: Creates a new Workbook, clears the RibbonXml property to hide the Ribbon UI, applies default OdsSaveOptions, and saves the file as an ODS document so the disabled ribbon setting is retained.
// Keywords: Aspose.Cells | C# | ODS | RibbonXml | disable ribbon UI | OdsSaveOptions | export to ODS | hide ribbon | OpenDocument Spreadsheet | workbook UI customization
// Common Searches: Aspose.Cells disable ribbon before ODS export | How to remove ribbon UI from ODS file using C# | Persist RibbonXml setting when saving as ODS | C# Aspose.Cells hide ribbon UI | Export Excel to ODS without ribbon
// Developer Intent: The developer needs to turn off the Ribbon UI in a workbook and ensure the setting is preserved when the file is saved as an ODS document.
// Use Cases: Generate server‑side ODS reports that open without a Ribbon interface. | Convert Excel workbooks to ODS while stripping UI elements for clean distribution. | Automate creation of ODS files for web apps where the Ribbon must be hidden in compatible editors.
// AI Prompts: Provide C# code that clears RibbonXml and saves a workbook as ODS using Aspose.Cells. | Show an example of disabling the Ribbon UI in an Aspose.Cells workbook and persisting the setting in the ODS output. | Explain how to verify that the Ribbon UI is disabled in the resulting ODS file after export.

using System;
using Aspose.Cells;
using Aspose.Cells.Ods;

// Creates a new Workbook, clears the RibbonXml property to hide the Ribbon UI, applies default OdsSaveOptions, and saves the file as an ODS document so the disabled ribbon setting is retained.
class Program
{
    static void Main()
    {
        // Create a new workbook instance
        Workbook workbook = new Workbook();

        // Disable the Ribbon UI by clearing the RibbonXml property
        workbook.RibbonXml = string.Empty;

        // Create ODS save options (default configuration)
        OdsSaveOptions odsOptions = new OdsSaveOptions();

        // Save the workbook as an ODS file; the disabled ribbon setting will be persisted
        workbook.Save("output.ods", odsOptions);
    }
}
