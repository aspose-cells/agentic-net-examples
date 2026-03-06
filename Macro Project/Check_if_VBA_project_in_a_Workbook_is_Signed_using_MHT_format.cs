using System;
using System.IO;
using Aspose.Cells;

class CheckVbaSignedMht
{
    static void Main()
    {
        // Path to the MHT file that contains the workbook
        string mhtPath = Path.Combine(Environment.CurrentDirectory, "sample.mht");

        if (!File.Exists(mhtPath))
        {
            Console.WriteLine($"File not found: {mhtPath}");
            return;
        }

        // Load the workbook from the MHT file
        Workbook workbook = new Workbook(mhtPath);

        // Get the VBA project from the workbook
        var vbaProject = workbook.VbaProject;

        // Determine whether the VBA project is signed
        if (vbaProject != null && vbaProject.IsSigned)
        {
            Console.WriteLine("VBA project is signed.");
            Console.WriteLine("Signature is valid: " + vbaProject.IsValidSigned);
        }
        else
        {
            Console.WriteLine("VBA project is not signed or does not exist.");
        }
    }
}