using System;
using Aspose.Cells;

class ProtectWorkbookStructure
{
    static void Main()
    {
        // Load the existing workbook
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Protect the workbook structure with a password
        workbook.Protect(ProtectionType.Structure, "myPassword123");

        // Save the protected workbook
        string outputPath = "protected_output.xlsx";
        workbook.Save(outputPath, SaveFormat.Xlsx);

        // Release resources
        workbook.Dispose();
    }
}