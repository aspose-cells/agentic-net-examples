using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class CheckVbaProtection
{
    static void Main()
    {
        // Path to the macro‑enabled workbook (XLSM) you want to inspect
        string inputPath = "sample.xlsm";

        // Load the workbook
        Workbook workbook = new Workbook(inputPath);

        // Determine whether the workbook contains a VBA project
        if (workbook.HasMacro)
        {
            // Access the VBA project
            VbaProject vbaProject = workbook.VbaProject;

            // Output protection status
            Console.WriteLine("VBA Project IsProtected: " + vbaProject.IsProtected);
            Console.WriteLine("VBA Project IsLockedForViewing: " + vbaProject.IslockedForViewing);
        }
        else
        {
            Console.WriteLine("The workbook does not contain a VBA project.");
        }

        // Save the workbook as a TAB‑delimited file (macros are not preserved in this format)
        string tabOutputPath = "output.tab";
        workbook.Save(tabOutputPath, SaveFormat.TabDelimited);
        Console.WriteLine("Workbook saved as TAB file: " + tabOutputPath);
    }
}