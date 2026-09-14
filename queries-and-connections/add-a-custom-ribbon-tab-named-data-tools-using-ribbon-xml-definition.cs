// Title: Create a custom "Data Tools" ribbon tab with a Refresh Data button in an Excel macro-enabled workbook using Aspose.Cells for .NET
// AI Prompts: Write C# code that builds Ribbon XML to add a "Data Tools" tab with a large "Refresh Data" button, assigns it to Workbook.RibbonXml, and saves the workbook as an XLSM file with Aspose.Cells. | Explain how to confirm that the custom ribbon tab was applied by reading the Workbook.RibbonXml property after the file is saved. | Show how to extend the Ribbon XML to add another group or button under the "Data Tools" tab using Aspose.Cells. | Provide a step‑by‑step guide for creating a macro‑enabled Excel file with custom UI via Aspose.Cells, including required namespaces and the correct SaveFormat.
// Common Searches: how to add a custom ribbon tab in Excel using Aspose.Cells .NET | Aspose.Cells RibbonXml property example for macro-enabled XLSM | C# code to create a Data Tools ribbon group with Refresh Data button in Excel workbook | setting custom UI on Excel workbook with Aspose.Cells and saving as XLSM | retrieve and inspect Ribbon XML from an Aspose.Cells workbook
// Tags: custom ribbon tab Aspose.Cells .NET | RibbonXml property macro-enabled XLSM | define Excel ribbon XML Aspose.Cells | add refresh data button to Excel ribbon | data tools group custom UI Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsRibbonDemo
{
    // The program creates a new Workbook, defines Ribbon XML that adds a "Data Tools" tab containing a large "Refresh Data" button, assigns the XML to the workbook via the RibbonXml property, and saves the file as a macro‑enabled XLSM workbook to preserve the custom UI.
    class Program
    {
        static void Main()
        {
            // Create a new workbook instance
            Workbook workbook = new Workbook();

            // Define the Ribbon XML that adds a custom tab named "Data Tools"
            string ribbonXml =
                "<customUI xmlns=\"http://schemas.microsoft.com/office/2006/01/customui\">" +
                "  <ribbon>" +
                "    <tabs>" +
                "      <tab id=\"dataToolsTab\" label=\"Data Tools\">" +
                "        <group id=\"dataToolsGroup\" label=\"Data Operations\">" +
                "          <button id=\"refreshButton\" label=\"Refresh Data\" size=\"large\" />" +
                "        </group>" +
                "      </tab>" +
                "    </tabs>" +
                "  </ribbon>" +
                "</customUI>";

            // Assign the Ribbon XML to the workbook
            workbook.RibbonXml = ribbonXml;

            // Save the workbook as a macro-enabled file (XLSM) to preserve the custom UI
            workbook.Save("DataToolsWorkbook.xlsm", SaveFormat.Xlsm);

            // Optional: confirm that the RibbonXml property has been set
            Console.WriteLine("Custom Ribbon tab 'Data Tools' has been added.");
        }
    }
}
