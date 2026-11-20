// Title: Apply password protection to Excel workbook view settings (structure and windows) using Aspose.Cells for .NET
// Common Searches: how to protect Excel workbook windows and structure with a password in C# using Aspose.Cells | C# Aspose.Cells example to lock zoom level and other view options | prevent users from changing default view settings in an .xlsx file with Aspose.Cells | Aspose.Cells protect workbook view settings programmatically in .NET | set workbook protection for structure and windows in Aspose.Cells C#
// Tags: Aspose.Cells workbook protection structure windows | C# password protect Excel view options | prevent zoom level modification Aspose.Cells | lock default view settings in .xlsx using Aspose | Aspose.Cells protect workbook windows API

using System;
using System.IO;
using Aspose.Cells;

// The C# program loads or creates an Excel workbook, then applies password‑protected protection for both the workbook's structure and its window settings, preventing changes to default view options such as zoom level, and saves the protected file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";
        const string password = "myPassword";

        Workbook workbook = null;

        try
        {
            // Verify that the input file exists; if not, create a new workbook.
            if (File.Exists(inputPath))
            {
                // Load the existing workbook.
                workbook = new Workbook(inputPath);
            }
            else
            {
                // Create a new workbook with a default worksheet.
                workbook = new Workbook();
                workbook.Worksheets[0].Name = "Sheet1";
            }

            // Protect the workbook's structure and windows with a password.
            workbook.Protect(ProtectionType.Structure | ProtectionType.Windows, password);

            // Save the protected workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors.
            Console.WriteLine($"Error: {ex.Message}");
        }
        finally
        {
            // Dispose the workbook if needed.
            workbook?.Dispose();
        }
    }
}
