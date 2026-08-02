using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class RetrieveVbaProject
{
    static void Main()
    {
        // Load an existing macro-enabled workbook
        string inputPath = "input.xlsm";
        Workbook workbook = new Workbook(inputPath);

        // Retrieve the VBA project from the workbook
        VbaProject vbaProject = workbook.VbaProject;

        // Verify that the VBA project exists and display some properties
        if (vbaProject != null)
        {
            Console.WriteLine("VBA Project Name: " + vbaProject.Name);
            Console.WriteLine("Is Protected: " + vbaProject.IsProtected);
            Console.WriteLine("Is Locked For Viewing: " + vbaProject.IslockedForViewing);
            Console.WriteLine("Is Signed: " + vbaProject.IsSigned);
        }
        else
        {
            Console.WriteLine("No VBA project found in the workbook.");
        }

        // Save the workbook (no modifications made) to demonstrate the save lifecycle
        string outputPath = "output.xlsm";
        workbook.Save(outputPath, SaveFormat.Xlsm);
    }
}