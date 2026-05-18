using System;
using Aspose.Cells;

class DisableDataTabDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Custom Ribbon XML that omits the default "Data" tab.
        // Only the desired built‑in tabs are re‑declared.
        string ribbonXml =
            "<customUI xmlns=\"http://schemas.microsoft.com/office/2006/01/customui\">" +
            "  <ribbon>" +
            "    <tabs>" +
            "      <tab idMso=\"TabHome\" />" +
            "      <tab idMso=\"TabInsert\" />" +
            "      <tab idMso=\"TabPageLayout\" />" +
            "      <tab idMso=\"TabFormulas\" />" +
            "      <tab idMso=\"TabReview\" />" +
            "      <tab idMso=\"TabView\" />" +
            "    </tabs>" +
            "  </ribbon>" +
            "</customUI>";

        // Apply the custom Ribbon definition to the workbook.
        workbook.RibbonXml = ribbonXml;

        // Save the workbook (macro‑enabled format preserves the Ribbon XML).
        workbook.Save("WorkbookWithoutDataTab.xlsm");
    }
}