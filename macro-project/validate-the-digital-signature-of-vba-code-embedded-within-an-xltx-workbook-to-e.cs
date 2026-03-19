using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

class ValidateVbaSignature
{
    static void Main()
    {
        const string sourcePath = "template_with_vba.xltx";
        const string copyPath = "copy_of_template.xltx";

        if (!File.Exists(sourcePath))
        {
            Console.WriteLine($"Source file not found: {sourcePath}");
            return;
        }

        // Load the workbook that contains a VBA project
        Workbook workbook = new Workbook(sourcePath);

        // Access the VBA project
        VbaProject vbaProject = workbook.VbaProject;

        // Verify if the VBA project is signed and if the signature is valid
        if (vbaProject != null && vbaProject.IsSigned)
        {
            Console.WriteLine("VBA project is signed.");
            Console.WriteLine("Signature is valid: " + vbaProject.IsValidSigned);
        }
        else
        {
            Console.WriteLine("VBA project is not signed.");
        }

        // Save a copy to ensure the signature information persists
        workbook.Save(copyPath, SaveFormat.Xltx);

        // Reload the saved copy and re‑check the signature status
        Workbook reloadedWorkbook = new Workbook(copyPath);
        VbaProject reloadedVba = reloadedWorkbook.VbaProject;

        Console.WriteLine("After reload - Is Signed: " + (reloadedVba != null && reloadedVba.IsSigned));
        Console.WriteLine("After reload - Is Valid Signed: " + (reloadedVba != null && reloadedVba.IsValidSigned));
    }
}