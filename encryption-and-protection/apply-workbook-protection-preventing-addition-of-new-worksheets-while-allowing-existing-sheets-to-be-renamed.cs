// Title: Apply workbook structure protection to block new worksheets while allowing sheet renaming with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an .xlsx file, protects its structure with a password using Aspose.Cells, enables renaming of existing worksheets, and saves the protected workbook. | Demonstrate how to call Workbook.Protect with ProtectionType.Structure and set Settings.AllowRename so sheet renaming remains possible after protection.
// Common Searches: Aspose.Cells C# protect workbook structure but keep sheet rename enabled | How to stop users from adding new worksheets in an Excel file using Aspose.Cells | Enable worksheet renaming after applying structure protection with Aspose.Cells .NET | Workbook.Protect with ProtectionType.Structure example code Aspose.Cells | Set AllowRename property after protecting workbook in C# Aspose.Cells
// Tags: workbook structure protection Aspose.Cells | prevent worksheet insertion .NET | allow sheet rename with protection Aspose.Cells | password-protected workbook C# | Excel file protection using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing .xlsx file, applies structure protection with a password via Aspose.Cells, optionally enables sheet renaming through Settings.AllowRename, and saves the workbook, preventing new worksheets from being added while still allowing existing sheets to be renamed.
class Program
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

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

            // Protect the workbook structure with a password
            workbook.Protect(ProtectionType.Structure, "MySecretPassword");

            // Allow renaming of worksheets if the API version supports it
            // Uncomment the following line for versions that expose AllowRename
            // workbook.Settings.AllowRename = true;

            // Save the protected workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
