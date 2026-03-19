using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class Program
{
    static void Main()
    {
        // Path to the XLSB workbook
        string filePath = "sample.xlsb";

        // Load the workbook (lifecycle rule: load)
        Workbook workbook = new Workbook(filePath);

        // Access the VBA project
        VbaProject vbaProject = workbook.VbaProject;

        // Verify that the workbook actually contains a VBA project
        if (workbook.HasMacro && vbaProject != null)
        {
            // Use VbaProject.IsProtected to determine if the project is protected
            Console.WriteLine("VBA Project Protected: " + vbaProject.IsProtected);

            // Use VbaProject.IslockedForViewing to determine if it is locked for viewing
            Console.WriteLine("VBA Project Locked for Viewing: " + vbaProject.IslockedForViewing);
        }
        else
        {
            Console.WriteLine("The workbook does not contain a VBA project.");
        }
    }
}