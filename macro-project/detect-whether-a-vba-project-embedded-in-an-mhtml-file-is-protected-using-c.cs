using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

class DetectVbaProtectionInMhtml
{
    static void Main(string[] args)
    {
        // Path to the MHTML file that may contain an embedded Excel workbook with VBA
        string mhtmlPath = args.Length > 0 ? args[0] : "sample.mht";

        if (!File.Exists(mhtmlPath))
        {
            Console.WriteLine($"File not found: {mhtmlPath}");
            return;
        }

        // Load the MHTML file into a Workbook object
        Workbook workbook = new Workbook(mhtmlPath);

        // Verify that the workbook actually contains a VBA project (macro)
        if (workbook.HasMacro && workbook.VbaProject != null)
        {
            // Use the VbaProject.IsProtected property to determine protection status
            bool isProtected = workbook.VbaProject.IsProtected;

            Console.WriteLine($"VBA project protection status: {(isProtected ? "Protected" : "Not protected")}");
        }
        else
        {
            Console.WriteLine("No VBA project found in the MHTML file.");
        }
    }
}