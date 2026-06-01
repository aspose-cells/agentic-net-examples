using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class VbaProjectAnalysis
{
    static void Main()
    {
        // Path to the macro-enabled workbook
        string filePath = "sample.xlsm";

        // Load the workbook from disk using the string constructor
        Workbook workbook = new Workbook(filePath);

        // Obtain the VbaProject object
        VbaProject vbaProject = workbook.VbaProject;

        // Check if a VBA project exists
        if (vbaProject != null)
        {
            // Output analysis information
            Console.WriteLine("VBA Project Name: " + vbaProject.Name);
            Console.WriteLine("Is Signed: " + vbaProject.IsSigned);
            Console.WriteLine("Is Protected: " + vbaProject.IsProtected);
            Console.WriteLine("Is Locked For Viewing: " + vbaProject.IslockedForViewing);
            Console.WriteLine("Modules Count: " + vbaProject.Modules.Count);
            Console.WriteLine("References Count: " + vbaProject.References.Count);
        }
        else
        {
            Console.WriteLine("No VBA project found in the workbook.");
        }

        // Clean up
        workbook.Dispose();
    }
}