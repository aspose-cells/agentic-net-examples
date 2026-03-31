using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class VbaProtectionCheck
{
    static void Main()
    {
        // Load an Excel file that may contain a VBA project (macro-enabled workbook)
        Workbook workbook = new Workbook("sample.xlsm");

        // Retrieve the VBA project associated with the workbook
        VbaProject vbaProject = workbook.VbaProject;

        // Determine if the VBA project is protected
        bool isProtected = vbaProject.IsProtected;

        // Determine if the VBA project is locked for viewing (additional protection detail)
        bool isLockedForViewing = vbaProject.IslockedForViewing;

        // Output the results
        Console.WriteLine("VBA Project Protected: " + isProtected);
        Console.WriteLine("VBA Project Locked for Viewing: " + isLockedForViewing);
    }
}