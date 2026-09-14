// Title: How to set the 2010 custom Ribbon XML namespace and ISO 29500 strict compliance for an Aspose.Cells workbook in C#
// AI Prompts: Write C# code that assigns a 2010 customUI XML string to Workbook.RibbonXml, sets Workbook.Settings.Compliance to OoxmlCompliance.Iso29500_2008_Strict, and saves the file as a macro‑enabled .xlsm. | Demonstrate how to create an empty Aspose.Cells workbook, update its ribbon XML for newer Office versions, enable strict OOXML compliance, and output verification messages to the console.
// Common Searches: Aspose.Cells C# change RibbonXml namespace to 2010 | Enable strict ISO 29500 compliance for Aspose.Cells workbook | Save workbook with custom ribbon UI as macro‑enabled xlsm using Aspose.Cells | Update custom ribbon XML for future Office compatibility in .NET
// Tags: set RibbonXml customUI 2010 Aspose.Cells | configure OoxmlCompliance Iso29500_2008_Strict | save macro-enabled .xlsm with custom ribbon | future‑compatible ribbon XML Office Aspose.Cells | Aspose.Cells workbook compliance mode

using System;
using Aspose.Cells;

namespace AsposeCellsRibbonUpdate
{
    // The example creates an empty Workbook, assigns RibbonXml using the 2010 customUI namespace, configures the workbook for ISO 29500 strict OOXML compliance, saves it as a macro‑enabled .xlsm file, and prints verification details to the console.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (empty workbook)
            Workbook workbook = new Workbook();

            // Update the Ribbon XML to use the newer custom UI namespace.
            // The 2010 namespace (http://schemas.microsoft.com/office/2010/06/customui)
            // ensures compatibility with newer Office versions while still being
            // understood by older versions that fall back to the 2006 namespace.
            string ribbonXml =
                "<customUI xmlns=\"http://schemas.microsoft.com/office/2010/06/customui\">" +
                "  <ribbon>" +
                "    <tabs>" +
                "      <tab id=\"customTab\" label=\"My Tab\">" +
                "        <group id=\"customGroup\" label=\"My Group\">" +
                "          <button id=\"customButton\" label=\"My Button\" size=\"large\" />" +
                "        </group>" +
                "      </tab>" +
                "    </tabs>" +
                "  </ribbon>" +
                "</customUI>";

            // Assign the updated Ribbon XML to the workbook
            workbook.RibbonXml = ribbonXml;

            // Set OOXML compliance to the strict ISO standard for future-proofing
            workbook.Settings.Compliance = OoxmlCompliance.Iso29500_2008_Strict;

            // Save the workbook as a macro-enabled file to retain the Ribbon UI
            workbook.Save("UpdatedRibbonWorkbook.xlsm");

            // Simple verification output
            Console.WriteLine("Ribbon XML set: " + (workbook.RibbonXml != null));
            Console.WriteLine("Compliance mode: " + workbook.Settings.Compliance);
            Console.WriteLine("Workbook saved as UpdatedRibbonWorkbook.xlsm");
        }
    }
}
