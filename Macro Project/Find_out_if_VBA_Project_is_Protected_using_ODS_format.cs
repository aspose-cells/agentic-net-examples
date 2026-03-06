using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class CheckVbaProtectionInOds
{
    static void Main()
    {
        // Path to the ODS file to be examined
        string odsFilePath = "sample.ods";

        // Load the ODS workbook
        Workbook workbook = new Workbook(odsFilePath);

        // Retrieve the VBA project associated with the workbook
        VbaProject vbaProject = workbook.VbaProject;

        // Check whether the VBA project is protected
        bool isVbaProtected = vbaProject.IsProtected;

        // Output the protection status
        Console.WriteLine($"Is VBA Project Protected: {isVbaProtected}");
    }
}