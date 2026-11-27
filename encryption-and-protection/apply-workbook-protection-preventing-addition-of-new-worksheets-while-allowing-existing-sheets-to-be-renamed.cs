// Title: Apply password‑protected structure protection to an Excel workbook in C# using Aspose.Cells, allowing sheet renaming but preventing new worksheet creation
// Common Searches: C# Aspose.Cells protect workbook structure with password to stop adding worksheets | How to allow sheet rename while disabling new sheet creation using Aspose.Cells .NET | Aspose.Cells workbook.Protect structure example in C# | Prevent users from inserting new worksheets in Excel via Aspose.Cells API | Set workbook structure protection without locking sheet names Aspose.Cells
// Tags: Aspose.Cells workbook structural lock | C# password protected Excel workbook | prevent new worksheet addition Aspose.Cells | rename sheets with protected workbook | Excel structure lock .NET

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing Excel file with Aspose.Cells, applies structure protection using a password so that new worksheets cannot be added while existing sheets can still be renamed, and saves the protected workbook, handling file‑not‑found and runtime errors.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";
        const string password = "MyPassword";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Protect the workbook structure with a password
            workbook.Protect(ProtectionType.Structure, password);

            // Save the protected workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors during processing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
