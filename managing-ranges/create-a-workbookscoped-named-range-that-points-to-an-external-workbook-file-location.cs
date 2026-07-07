using System;
using Aspose.Cells;

namespace AsposeCellsExternalNamedRange
{
    class Program
    {
        static void Main()
        {
            // Create the main workbook
            Workbook mainWb = new Workbook();

            // Define external workbook file name and the sheet(s) it contains
            string externalFileName = "ExternalWorkbook.xlsx";
            string[] externalSheets = new string[] { "Sheet1" };

            // Add an external link to the main workbook
            int linkIndex = mainWb.Worksheets.ExternalLinks.Add(externalFileName, externalSheets);
            ExternalLink externalLink = mainWb.Worksheets.ExternalLinks[linkIndex];

            // Add an external name that refers to a range inside the external workbook
            // The range reference is relative to the external sheet
            externalLink.AddExternalName("ExtRange", "Sheet1!$A$1:$B$2");

            // Create a workbook‑scoped named range that points to the external name
            // Workbook‑scoped names have SheetIndex = 0 (global)
            int nameIndex = mainWb.Worksheets.Names.Add("MyExternalRange");
            Name workbookScopedName = mainWb.Worksheets.Names[nameIndex];
            // The RefersTo formula uses the external file name and the external name
            workbookScopedName.RefersTo = $"=[{externalFileName}]!ExtRange";

            // Optionally, demonstrate usage of the named range in a formula
            Worksheet sheet = mainWb.Worksheets[0];
            sheet.Cells["A1"].Formula = "=SUM(MyExternalRange)";

            // Save the workbook
            mainWb.Save("WorkbookWithExternalNamedRange.xlsx");
        }
    }
}