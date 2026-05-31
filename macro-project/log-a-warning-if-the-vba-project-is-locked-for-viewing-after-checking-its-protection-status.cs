using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class CheckVbaProjectLock
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsm");

        // Get the VBA project associated with the workbook
        VbaProject vbaProject = workbook.VbaProject;

        // Determine if the VBA project is protected and if it is locked for viewing
        bool isProtected = vbaProject.IsProtected;
        bool isLockedForViewing = vbaProject.IslockedForViewing;

        // Log a warning when the project is both protected and locked for viewing
        if (isProtected && isLockedForViewing)
        {
            Console.WriteLine("Warning: VBA project is protected and locked for viewing.");
        }
        else
        {
            Console.WriteLine($"VBA project status - Protected: {isProtected}, LockedForViewing: {isLockedForViewing}");
        }

        // Save the workbook (no modifications made, just to follow lifecycle)
        workbook.Save("output.xlsm", SaveFormat.Xlsm);
    }
}