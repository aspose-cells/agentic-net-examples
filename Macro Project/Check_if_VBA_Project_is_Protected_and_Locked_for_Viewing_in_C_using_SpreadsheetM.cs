using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class CheckVbaProtection
{
    static void Main()
    {
        // Path to the macro‑enabled workbook (XLSM)
        string inputPath = "sample.xlsm";

        // Load the workbook (preserves VBA project)
        Workbook workbook = new Workbook(inputPath);

        // Get the VBA project associated with the workbook
        VbaProject vbaProject = workbook.VbaProject;

        // Determine whether the VBA project is protected
        bool isProtected = vbaProject.IsProtected;

        // Determine whether the VBA project is locked for viewing
        bool isLockedForViewing = vbaProject.IslockedForViewing;

        // Output the results
        Console.WriteLine($"VBA Project Protected: {isProtected}");
        Console.WriteLine($"VBA Project Locked for Viewing: {isLockedForViewing}");

        // Save the workbook in SpreadsheetML (XML) format if needed
        string outputPath = "sample.xml";
        workbook.Save(outputPath, SaveFormat.Xml);
    }
}