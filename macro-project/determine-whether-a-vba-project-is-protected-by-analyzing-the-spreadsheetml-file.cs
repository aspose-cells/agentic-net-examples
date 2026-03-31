using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class Program
{
    static void Main()
    {
        // Path to the macro-enabled Excel file to be analyzed
        string filePath = "input.xlsm";

        // Load the workbook from the specified file
        Workbook workbook = new Workbook(filePath);

        // Retrieve the VBA project associated with the workbook
        VbaProject vbaProject = workbook.VbaProject;

        // Check if the VBA project is protected
        bool isProtected = vbaProject.IsProtected;

        // Check if the VBA project is locked for viewing (optional additional info)
        bool isLockedForViewing = vbaProject.IslockedForViewing;

        // Output the protection status
        Console.WriteLine($"VBA Project Protected: {isProtected}");
        Console.WriteLine($"VBA Project Locked for Viewing: {isLockedForViewing}");
    }
}