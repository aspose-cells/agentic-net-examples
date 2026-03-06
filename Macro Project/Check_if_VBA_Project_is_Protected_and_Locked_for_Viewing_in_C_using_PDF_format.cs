using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class CheckVbaProtection
{
    static void Main()
    {
        // Load a macro-enabled Excel workbook (XLSM) that may contain a VBA project
        Workbook workbook = new Workbook("input.xlsm");

        // Access the VBA project associated with the workbook
        VbaProject vbaProject = workbook.VbaProject;

        // Determine whether the VBA project is protected
        bool isProtected = vbaProject.IsProtected;

        // Determine whether the VBA project is locked for viewing
        bool isLockedForViewing = vbaProject.IslockedForViewing;

        // Output the protection status to the console
        Console.WriteLine($"VBA Project Protected: {isProtected}");
        Console.WriteLine($"VBA Project Locked for Viewing: {isLockedForViewing}");

        // Save the workbook as a PDF file (the VBA protection status does not affect PDF conversion)
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}