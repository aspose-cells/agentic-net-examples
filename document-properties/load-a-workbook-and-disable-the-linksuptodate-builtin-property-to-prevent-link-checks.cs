using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Disable the LinksUpToDate built‑in property to prevent link checks
        workbook.BuiltInDocumentProperties.LinksUpToDate = false;

        // Save the workbook with the updated setting
        workbook.Save("output.xlsx");
    }
}