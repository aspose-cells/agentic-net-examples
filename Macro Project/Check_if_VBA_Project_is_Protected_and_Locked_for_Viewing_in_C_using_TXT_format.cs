using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

class CheckVbaProtection
{
    static void Main(string[] args)
    {
        // Input Excel file (macro-enabled) path; default if not provided
        string inputPath = args.Length > 0 ? args[0] : "input.xlsm";

        // Output text file path; default if not provided
        string outputPath = args.Length > 1 ? args[1] : "VbaProtectionInfo.txt";

        // Load the workbook
        Workbook workbook = new Workbook(inputPath);

        // Access the VBA project
        VbaProject vbaProject = workbook.VbaProject;

        // Retrieve protection information
        bool isProtected = vbaProject.IsProtected;
        bool isLockedForViewing = vbaProject.IslockedForViewing;

        // Prepare result lines
        string[] resultLines = new string[]
        {
            $"File: {Path.GetFileName(inputPath)}",
            $"VBA Project Protected: {isProtected}",
            $"VBA Project Locked for Viewing: {isLockedForViewing}"
        };

        // Write results to a text file
        File.WriteAllLines(outputPath, resultLines);

        // Also output to console
        foreach (string line in resultLines)
        {
            Console.WriteLine(line);
        }
    }
}