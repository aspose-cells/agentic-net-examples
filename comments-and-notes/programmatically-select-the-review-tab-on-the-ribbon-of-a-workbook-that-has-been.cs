using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the existing XLSX workbook
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Define Ribbon XML that selects the built‑in Review tab when the workbook is opened
        string ribbonXml =
            @"<customUI xmlns=""http://schemas.microsoft.com/office/2006/01/customui"">" +
            "  <ribbon>" +
            "    <tabs>" +
            "      <tab idMso=\"ReviewTab\"/>" +
            "    </tabs>" +
            "  </ribbon>" +
            "</customUI>";

        // Apply the custom Ribbon XML to the workbook
        workbook.RibbonXml = ribbonXml;

        // Save the workbook (the Review tab will be selected on open)
        string outputPath = "output.xlsx";
        workbook.Save(outputPath, SaveFormat.Xlsx);
    }
}