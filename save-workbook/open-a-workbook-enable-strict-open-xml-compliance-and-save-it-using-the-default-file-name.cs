using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (uses the Workbook() constructor rule)
        Workbook workbook = new Workbook();

        // Enable strict Open XML compliance (uses WorkbookSettings.Compliance property)
        workbook.Settings.Compliance = OoxmlCompliance.Iso29500_2008_Strict;

        // Define a default file name and save the workbook (uses Workbook.Save(string) method)
        workbook.FileName = "StrictComplianceWorkbook.xlsx";
        workbook.Save(workbook.FileName);
    }
}