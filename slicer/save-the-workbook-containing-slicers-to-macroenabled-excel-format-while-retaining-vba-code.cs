using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load an existing workbook that contains slicers (and possibly VBA code)
        Workbook workbook = new Workbook("InputWithSlicers.xlsx");

        // Save the workbook as a macro‑enabled file (XLSM) to retain any VBA/macros
        workbook.Save("OutputWithMacros.xlsm", SaveFormat.Xlsm);
    }
}