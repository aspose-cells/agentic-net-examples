using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Author: Aspose.Cells .NET example
        // Create LoadOptions and enable ignoring of useless shapes (including pictures) to improve load performance.
        LoadOptions loadOptions = new LoadOptions
        {
            IgnoreUselessShapes = true
        };

        // Load the workbook with the specified options.
        Workbook workbook = new Workbook("LargeWorkbook.xlsx", loadOptions);

        // Save the workbook (optional, demonstrates that loading succeeded without pictures).
        workbook.Save("LargeWorkbook_NoPictures.xlsx");
    }
}