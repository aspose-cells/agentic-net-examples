using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class CheckVbaProtection
{
    static void Main()
    {
        // Path to the MHTML workbook
        string filePath = "workbook.mht";

        // Load the workbook (Aspose.Cells automatically detects MHTML format)
        Workbook workbook = new Workbook(filePath);

        // Verify that the workbook actually contains a VBA project
        if (!workbook.HasMacro)
        {
            Console.WriteLine("The workbook does not contain any VBA project.");
            return;
        }

        // Access the VBA project
        VbaProject vbaProject = workbook.VbaProject;

        // Determine if the VBA project is password‑protected
        bool isProtected = vbaProject.IsProtected;

        // Determine if the VBA project is locked for viewing
        bool isLockedForViewing = vbaProject.IslockedForViewing;

        // Output the results
        Console.WriteLine($"VBA Project Protected: {isProtected}");
        Console.WriteLine($"VBA Project Locked for Viewing: {isLockedForViewing}");
    }
}