using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

class Program
{
    static void Main()
    {
        // Path to the XLSM workbook that may contain a VBA project
        string workbookFileName = "sample_with_vba_signed.xlsm";
        string workbookPath = Path.Combine(Directory.GetCurrentDirectory(), workbookFileName);

        if (!File.Exists(workbookPath))
        {
            Console.WriteLine($"Workbook file not found: {workbookPath}");
            return;
        }

        // Load the workbook
        Workbook workbook = new Workbook(workbookPath);

        // Access the VBA project associated with the workbook
        VbaProject vbaProject = workbook.VbaProject;

        // Check if the VBA project is signed
        bool isSigned = vbaProject.IsSigned;
        Console.WriteLine("VBA Project Signed: " + isSigned);

        // If signed, verify whether the signature is valid
        if (isSigned)
        {
            bool isValid = vbaProject.IsValidSigned;
            Console.WriteLine("VBA Project Signature Valid: " + isValid);
        }
        else
        {
            Console.WriteLine("VBA Project is not signed, so no signature validation is performed.");
        }
    }
}