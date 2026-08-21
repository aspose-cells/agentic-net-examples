// Title: Import External Ribbon XML into an Existing Workbook with Aspose.Cells for .NET
// Description: Loads an existing Excel file, reads a custom ribbon definition from an external .xml file, assigns it to Workbook.RibbonXml, and saves the workbook as a macro‑enabled .xlsm to retain the custom UI.
// Keywords: Aspose.Cells | .NET | RibbonXml | custom ribbon | import ribbon xml | macro-enabled workbook | Excel UI customization | programmatic ribbon | external xml file | Workbook.RibbonXml property
// Common Searches: Aspose.Cells set custom ribbon XML | How to add custom ribbon to Excel using C# | Load ribbon XML from file with Aspose.Cells | Save workbook with custom ribbon as xlsm | Programmatically modify Excel ribbon UI .NET
// Developer Intent: The developer wants to embed a custom ribbon UI, defined in an external XML file, into an existing Excel workbook programmatically using Aspose.Cells.
// Use Cases: Apply a company‑wide ribbon layout to generated reports before distribution. | Attach a custom ribbon to a template workbook for macro‑enabled add‑ins. | Batch‑update ribbon definitions across multiple workbooks in an automated process.
// AI Prompts: Show C# code that reads a ribbon XML file and assigns it to Workbook.RibbonXml using Aspose.Cells, then saves the file as .xlsm. | Provide robust error handling for missing or malformed ribbon XML when embedding it with Aspose.Cells. | Explain how to verify that the custom ribbon appears correctly when opening the saved workbook in Excel. | Suggest ways to programmatically replace an existing ribbon definition with a new XML file for multiple workbooks.

using System;
using System.IO;
using Aspose.Cells;

// Loads an existing Excel file, reads a custom ribbon definition from an external .xml file, assigns it to Workbook.RibbonXml, and saves the workbook as a macro‑enabled .xlsm to retain the custom UI.
public class RibbonXmlDemo
{
    public static void Main()
    {
        // Path to the existing workbook
        string workbookPath = "input.xlsx";

        // Path to the external ribbon XML file
        string ribbonXmlPath = "customRibbon.xml";

        // Load the existing workbook
        Workbook workbook = new Workbook(workbookPath);

        // Read the ribbon XML content from the file
        string ribbonXml = File.ReadAllText(ribbonXmlPath);

        // Assign the ribbon XML to the workbook
        workbook.RibbonXml = ribbonXml;

        // Save the workbook (saving as macro-enabled to preserve the ribbon UI)
        workbook.Save("output.xlsm");

        Console.WriteLine("Custom ribbon XML has been applied and the workbook saved as 'output.xlsm'.");
    }
}
