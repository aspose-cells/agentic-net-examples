using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class CheckVbaProtection
{
    static void Main()
    {
        // Path to the XLSX file to be examined
        string filePath = "sample.xlsx";

        // Load the workbook (XLSX format)
        Workbook workbook = new Workbook(filePath);

        // Retrieve the VBA project associated with the workbook
        VbaProject vbaProject = workbook.VbaProject;

        // If there is no VBA project, inform the user
        if (vbaProject == null)
        {
            Console.WriteLine("No VBA project found in the workbook.");
            return;
        }

        // Output the protection status of the VBA project
        Console.WriteLine("Is VBA Project Protected: " + vbaProject.IsProtected);
        Console.WriteLine("Is VBA Project Locked for Viewing: " + vbaProject.IslockedForViewing);
    }
}