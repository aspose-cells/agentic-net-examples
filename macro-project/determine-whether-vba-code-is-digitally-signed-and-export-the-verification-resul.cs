using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

class VbaSignatureChecker
{
    static void Main(string[] args)
    {
        // Expect two arguments: input workbook path and output CSV path
        if (args.Length < 2)
        {
            Console.WriteLine("Usage: VbaSignatureChecker <inputWorkbookPath> <outputCsvPath>");
            return;
        }

        string workbookPath = args[0];
        string csvPath = args[1];

        // Load the workbook (lifecycle: load)
        Workbook workbook = new Workbook(workbookPath);

        // Access the VBA project
        VbaProject vbaProject = workbook.VbaProject;

        // Determine if VBA code is signed and if the signature is valid
        bool isSigned = vbaProject.IsSigned;
        bool isValidSigned = vbaProject.IsValidSigned;

        // Export results to CSV
        using (StreamWriter writer = new StreamWriter(csvPath, false))
        {
            writer.WriteLine("FileName,IsSigned,IsValidSigned");
            writer.WriteLine($"{Path.GetFileName(workbookPath)},{isSigned},{isValidSigned}");
        }

        Console.WriteLine($"Verification results saved to: {csvPath}");
    }
}