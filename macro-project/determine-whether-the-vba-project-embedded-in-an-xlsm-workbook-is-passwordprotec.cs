using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class CheckVbaProtection
{
    static void Main(string[] args)
    {
        // Path to the macro-enabled workbook (XLSM)
        string filePath = "sample.xlsm";
        if (args.Length > 0)
        {
            filePath = args[0];
        }

        // Load the workbook
        Workbook workbook = new Workbook(filePath);

        // Get the VBA project associated with the workbook
        VbaProject vbaProject = workbook.VbaProject;

        // Check if the VBA project is password‑protected (cannot be edited)
        bool isProtected = vbaProject.IsProtected;

        // Optional: also check if it is locked for viewing
        bool isLockedForViewing = vbaProject.IslockedForViewing;

        Console.WriteLine($"VBA Project Protected: {isProtected}");
        Console.WriteLine($"VBA Project Locked for Viewing: {isLockedForViewing}");
    }
}