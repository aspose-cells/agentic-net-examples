// Title: Protect an XLS workbook's structure with a complex password using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells protect workbook structure with password for XLS file | How to lock sheet order in an Excel .xls using Aspose.Cells .NET | Apply complex password to workbook structure protection in Aspose.Cells C# example | Prevent users from moving or deleting sheets in an XLS workbook with Aspose.Cells | Aspose.Cells Workbook.Protect structure password syntax C#
// Tags: Workbook.Protect structure XLS C# | Aspose.Cells workbook structure lock | complex password protection .NET Excel | disable sheet reorder Aspose.Cells | XLS workbook protection example

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The C# program loads an existing XLS workbook, applies structure protection with a complex password via Workbook.Protect, and saves the protected file to prevent sheet reordering or deletion.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xls";
            string outputPath = "output.xls";

            try
            {
                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Protect workbook structure with a complex password
                workbook.Protect(ProtectionType.Structure, "C0mpl3xP@ssw0rd!#2026");

                // Save the protected workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
