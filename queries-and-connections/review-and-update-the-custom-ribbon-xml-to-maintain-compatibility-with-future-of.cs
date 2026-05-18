using System;
using Aspose.Cells;

namespace AsposeCellsRibbonUpdateDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Define ribbon XML that supports both legacy (customUI) and newer (customUI14) schemas.
            // This ensures the ribbon customizations work in older Office versions (2007) and
            // future versions (Office 2010, 2013, 2016, 2019, 2021, etc.).
            string ribbonXml =
                // Legacy schema (Office 2007)
                "<customUI xmlns=\"http://schemas.microsoft.com/office/2006/01/customui\">" +
                "  <ribbon>" +
                "    <tabs>" +
                "      <tab idMso=\"TabHome\" label=\"Home\">" +
                "        <group id=\"customGroupLegacy\" label=\"Legacy Group\">" +
                "          <button id=\"legacyButton\" label=\"Legacy Button\" size=\"large\" onAction=\"OnLegacyButtonClick\" />" +
                "        </group>" +
                "      </tab>" +
                "    </tabs>" +
                "  </ribbon>" +
                "</customUI>" +

                // Newer schema (Office 2010+). The namespace version 2009/2010 is used for
                // Office 2010 and later, and Office 2016+ will also recognize it.
                "<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\" " +
                "          xmlns:bt=\"http://schemas.microsoft.com/office/2006/01/customui\">" +
                "  <ribbon>" +
                "    <tabs>" +
                "      <tab idMso=\"TabHome\" label=\"Home\">" +
                "        <group id=\"customGroupModern\" label=\"Modern Group\">" +
                "          <button id=\"modernButton\" label=\"Modern Button\" size=\"large\" onAction=\"OnModernButtonClick\" />" +
                "        </group>" +
                "      </tab>" +
                "    </tabs>" +
                "  </ribbon>" +
                "</customUI>";

            // Assign the combined XML to the workbook's RibbonXml property.
            workbook.RibbonXml = ribbonXml;

            // Optional: set OOXML compliance to the strict ISO standard for future‑proofing.
            workbook.Settings.Compliance = OoxmlCompliance.Iso29500_2008_Strict;

            // Save the workbook (lifecycle: save). Use .xlsm to preserve macros and ribbon XML.
            workbook.Save("CustomRibbonUpdated.xlsm");

            // Confirmation output.
            Console.WriteLine("Workbook saved with updated Ribbon XML.");
        }
    }
}