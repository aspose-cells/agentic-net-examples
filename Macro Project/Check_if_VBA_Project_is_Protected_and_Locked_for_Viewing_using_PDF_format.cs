using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class CheckVbaProtection
{
    static void Main()
    {
        // Path to the macro-enabled Excel file
        string excelPath = "sample.xlsm";

        // Load the workbook
        Workbook workbook = new Workbook(excelPath);

        // Get the VBA project associated with the workbook
        VbaProject vbaProject = workbook.VbaProject;

        // Determine if the VBA project is protected
        bool isProtected = vbaProject.IsProtected;

        // Determine if the VBA project is locked for viewing
        bool isLockedForViewing = vbaProject.IslockedForViewing;

        // Output the results
        Console.WriteLine($"VBA Project Protected: {isProtected}");
        Console.WriteLine($"VBA Project Locked for Viewing: {isLockedForViewing}");

        // Save the workbook as PDF (demonstrates using PDF format)
        string pdfOutput = "output.pdf";
        workbook.Save(pdfOutput, SaveFormat.Pdf);
    }
}