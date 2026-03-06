using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the protected workbook
        string inputPath = "protected_workbook.xlsx";
        // Password used to protect the workbook structure
        string password = "myPassword";

        // Load the workbook
        Workbook workbook = new Workbook(inputPath);

        // Unprotect the workbook structure
        workbook.Unprotect(password);

        // Save the unprotected workbook
        string outputPath = "unprotected_workbook.xlsx";
        workbook.Save(outputPath);
    }
}