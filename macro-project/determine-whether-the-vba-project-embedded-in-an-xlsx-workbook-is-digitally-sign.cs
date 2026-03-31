using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class CheckVbaSignature
{
    static void Main()
    {
        // Path to the workbook that may contain a VBA project (e.g., .xlsm)
        string workbookPath = "sample.xlsm";

        // Load the workbook from file
        Workbook workbook = new Workbook(workbookPath);

        // Get the VBA project associated with the workbook
        VbaProject vbaProject = workbook.VbaProject;

        // Determine if the VBA project is signed and whether the signature is valid
        if (vbaProject != null && vbaProject.IsSigned)
        {
            Console.WriteLine("VBA project is signed.");
            Console.WriteLine("Is the signature valid? " + vbaProject.IsValidSigned);
        }
        else
        {
            Console.WriteLine("VBA project is not signed.");
        }

        // Optionally, also check if the entire workbook is digitally signed
        Console.WriteLine("Workbook is digitally signed: " + workbook.IsDigitallySigned);
    }
}