using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class CheckVbaProtection
{
    static void Main()
    {
        // Load a macro-enabled workbook (XLSM)
        Workbook workbook = new Workbook("sample.xlsm");

        // Access the VBA project associated with the workbook
        VbaProject vbaProject = workbook.VbaProject;

        // Output protection status
        Console.WriteLine("Is VBA Project Protected: " + vbaProject.IsProtected);
        Console.WriteLine("Is VBA Project Locked for Viewing: " + vbaProject.IslockedForViewing);
    }
}