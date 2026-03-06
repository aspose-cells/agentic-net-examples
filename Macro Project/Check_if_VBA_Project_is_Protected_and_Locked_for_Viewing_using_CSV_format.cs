using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

class CheckVbaProtectionCsv
{
    static void Main(string[] args)
    {
        // Path to the macro-enabled workbook to inspect
        string inputPath = "input.xlsm";

        // Path where the CSV result will be saved
        string csvPath = "VbaProtectionStatus.csv";

        // Load the workbook
        Workbook workbook = new Workbook(inputPath);

        // Access the VBA project associated with the workbook
        VbaProject vbaProject = workbook.VbaProject;

        // Retrieve protection information
        bool isProtected = vbaProject.IsProtected;
        bool isLockedForViewing = vbaProject.IslockedForViewing;

        // Write the information to a CSV file
        using (StreamWriter writer = new StreamWriter(csvPath, false))
        {
            // CSV header
            writer.WriteLine("Property,Value");
            // Protection status rows
            writer.WriteLine($"IsProtected,{isProtected}");
            writer.WriteLine($"IsLockedForViewing,{isLockedForViewing}");
        }

        Console.WriteLine($"VBA protection status has been written to '{csvPath}'.");
    }
}