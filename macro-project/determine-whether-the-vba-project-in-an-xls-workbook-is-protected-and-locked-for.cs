using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class VbaProjectProtectionCheck
{
    static void Main(string[] args)
    {
        // Verify that a file path was provided
        if (args.Length == 0)
        {
            Console.WriteLine("Usage: VbaProjectProtectionCheck <path-to-workbook>");
            return;
        }

        string filePath = args[0];

        // Load the workbook (supports .xls, .xlsx, .xlsm, etc.)
        Workbook workbook = new Workbook(filePath);

        // Access the VBA project associated with the workbook
        VbaProject vbaProject = workbook.VbaProject;

        // Determine whether the VBA project is protected
        bool isProtected = vbaProject.IsProtected;

        // Determine whether the VBA project is locked for viewing
        bool isLockedForViewing = vbaProject.IslockedForViewing;

        // Output the results
        Console.WriteLine($"VBA Project Protected: {isProtected}");
        Console.WriteLine($"VBA Project Locked for Viewing: {isLockedForViewing}");
    }
}