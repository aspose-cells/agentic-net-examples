using System;
using Aspose.Cells;

class ProtectWorkbookStructure
{
    static void Main()
    {
        // Path to the source XLSX file
        string inputPath = "input.xlsx";

        // Load the workbook from the file
        Workbook workbook = new Workbook(inputPath);

        // Protect only the workbook structure with a password
        // This prevents adding, deleting, renaming, or moving worksheets
        workbook.Protect(ProtectionType.Structure, "StrongPassword123");

        // Save the protected workbook to a new file
        string outputPath = "protected_workbook.xlsx";
        workbook.Save(outputPath, SaveFormat.Xlsx);

        // Release resources
        workbook.Dispose();
    }
}