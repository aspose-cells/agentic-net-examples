using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class Program
{
    static void Main(string[] args)
    {
        // Output TSV header
        Console.WriteLine("FilePath\tIsProtected\tIsLockedForViewing");

        if (args.Length == 0)
        {
            Console.Error.WriteLine("Provide at least one Excel file path as a command‑line argument.");
            return;
        }

        foreach (string filePath in args)
        {
            try
            {
                // Load the workbook (macro‑enabled or not)
                Workbook workbook = new Workbook(filePath);

                // Access the VBA project associated with the workbook
                VbaProject vbaProject = workbook.VbaProject;

                // Determine protection status
                bool isProtected = vbaProject.IsProtected;
                bool isLockedForViewing = vbaProject.IslockedForViewing;

                // Output results in TSV format
                Console.WriteLine($"{filePath}\t{isProtected}\t{isLockedForViewing}");
            }
            catch (Exception ex)
            {
                // Report any errors while processing the file
                Console.Error.WriteLine($"Error processing '{filePath}': {ex.Message}");
            }
        }
    }
}