using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create the main workbook
        Workbook workbook = new Workbook();

        // External workbook file and sheet name
        string externalFileName = "ExternalData.xlsx";
        string externalSheetName = "Sheet1";

        // Add an external link entry (optional but records the link)
        ExternalLinkCollection externalLinks = workbook.Worksheets.ExternalLinks;
        int linkIdx = externalLinks.Add(externalFileName, new string[] { externalSheetName });

        // Add a workbook‑scoped named range
        int nameIdx = workbook.Worksheets.Names.Add("ExternalRange");
        Name externalRangeName = workbook.Worksheets.Names[nameIdx];

        // Point the name to a range in the external workbook (e.g., A1:C3 on Sheet1)
        externalRangeName.RefersTo = $"='[{externalFileName}]{externalSheetName}'!$A$1:$C$3";

        // Example usage of the named range in a formula
        workbook.Worksheets[0].Cells["A1"].Formula = "=SUM(ExternalRange)";

        // Calculate formulas (will succeed if the external file exists)
        workbook.CalculateFormula();

        // Save the workbook
        workbook.Save("WorkbookWithExternalNamedRange.xlsx");
    }
}