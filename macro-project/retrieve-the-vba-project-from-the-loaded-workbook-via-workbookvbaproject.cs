using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class Program
{
    static void Main()
    {
        // Load an existing macro-enabled workbook
        Workbook workbook = new Workbook("input.xlsm");

        // Retrieve the VBA project from the workbook
        VbaProject vbaProject = workbook.VbaProject;

        // Display some information about the VBA project
        if (vbaProject != null)
        {
            Console.WriteLine("VBA Project Name: " + vbaProject.Name);
            Console.WriteLine("Is Signed: " + vbaProject.IsSigned);
            Console.WriteLine("Is Protected: " + vbaProject.IsProtected);
            Console.WriteLine("Is Locked For Viewing: " + vbaProject.IslockedForViewing);
        }
        else
        {
            Console.WriteLine("The workbook does not contain a VBA project.");
        }
    }
}