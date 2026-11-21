// Title: Protect workbook structure only while keeping cells editable with Aspose.Cells for .NET
// Common Searches: asp.net protect workbook structure only using Aspose.Cells C# | allow users to edit cells on unlocked sheets while workbook is protected Aspose.Cells | protect Excel workbook structure programmatically with a password using Aspose.Cells .NET | how to lock workbook structure but keep cell content editable Aspose.Cells
// Tags: Workbook.Protect structure password C# | Aspose.Cells workbook structure lock | editable cells on unprotected sheets Aspose.Cells | Excel file structure protection .NET | prevent sheet addition removal Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads or creates an Excel workbook, applies structure‑only protection with a password via Workbook.Protect(ProtectionType.Structure, "StrongPassword123"), and saves the file, allowing users to edit cells on sheets that are not locked.
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";

                Workbook workbook;

                // Load existing workbook if it exists; otherwise create a new one.
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    workbook = new Workbook();
                }

                // Protect only the workbook structure with a password.
                workbook.Protect(ProtectionType.Structure, "StrongPassword123");

                // Save the protected workbook.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
