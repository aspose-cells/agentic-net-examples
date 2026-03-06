using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class CheckVbaProtection
{
    static void Main(string[] args)
    {
        // Verify that a file path was provided
        if (args.Length == 0)
        {
            Console.WriteLine("Usage: CheckVbaProtection <excel_file_path>");
            return;
        }

        string filePath = args[0];

        // Load the workbook (must be a macro-enabled file to contain a VBA project)
        Workbook workbook = new Workbook(filePath);

        // Access the VBA project associated with the workbook
        VbaProject vbaProject = workbook.VbaProject;

        // Retrieve protection information
        bool isProtected = vbaProject.IsProtected;
        bool isLockedForViewing = vbaProject.IslockedForViewing;

        // Output the results in TSV format: FilePath    IsProtected    IsLockedForViewing
        Console.WriteLine("FilePath\tIsProtected\tIsLockedForViewing");
        Console.WriteLine($"{filePath}\t{isProtected}\t{isLockedForViewing}");
    }
}