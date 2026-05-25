using System;
using Aspose.Cells;

class RemoveExternalLinksDemo
{
    static void Main()
    {
        // Load the workbook that may contain external links
        Workbook workbook = new Workbook("input.xlsx");

        // Remove all external links.
        // The boolean parameter 'true' updates formulas to refer to the current workbook
        // wherever possible, otherwise the formulas would be cleared.
        workbook.Worksheets.ExternalLinks.Clear(true);

        // Save the cleaned workbook ready for distribution
        workbook.Save("output.xlsx");
    }
}