// Title: Protect only the workbook structure while keeping cells editable on unlocked worksheets using Aspose.Cells for .NET
// Common Searches: Aspose.Cells protect workbook structure without password in C# | Allow cell editing on unlocked sheets while workbook structure is locked using Aspose.Cells | C# code to prevent adding or deleting worksheets but keep cell editing enabled | How to apply structure‑only protection to an Excel file with Aspose.Cells | Save Excel workbook with structure protection and no password using Aspose.Cells .NET
// Tags: Workbook.Protect structure protection Aspose.Cells | Excel workbook structure lock C# | Editable cells on unlocked worksheets Aspose.Cells | Prevent sheet addition deletion Aspose.Cells | Save protected workbook without password .NET

using System;
using System.IO;
using Aspose.Cells;

// Loads an existing Excel file, applies structure‑only protection (preventing sheet add, delete, rename, or move) with an empty password via Workbook.Protect, and saves the protected workbook while leaving cells on unlocked sheets editable.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Protect only the workbook structure (prevents adding, deleting, renaming, or moving worksheets)
            // An empty password is supplied because the overload requires it.
            workbook.Protect(ProtectionType.Structure, string.Empty);

            // Save the workbook with the applied protection
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
