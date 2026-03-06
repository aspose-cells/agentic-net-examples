using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class CheckVbaProtection
{
    static void Main()
    {
        // Path to the Excel file (macro‑enabled .xlsm)
        string filePath = "sample.xlsm";

        // Load the workbook
        Workbook wb = new Workbook(filePath);

        // Get the VBA project associated with the workbook
        VbaProject vbaProject = wb.VbaProject;

        // Determine whether the workbook actually contains VBA modules
        bool hasVba = vbaProject.Modules.Count > 0;

        if (hasVba)
        {
            Console.WriteLine("VBA project found.");
            // Use the IsProtected property to check protection status
            Console.WriteLine("Is VBA Project Protected: " + vbaProject.IsProtected);
            // Additional info: locked for viewing
            Console.WriteLine("Is VBA Project Locked for Viewing: " + vbaProject.IslockedForViewing);
        }
        else
        {
            Console.WriteLine("No VBA project present in the workbook.");
        }

        // Clean up
        wb.Dispose();
    }
}