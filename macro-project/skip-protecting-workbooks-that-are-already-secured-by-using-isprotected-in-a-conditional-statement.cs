// Title: Use Aspose.Cells IsProtected to conditionally protect an Excel workbook in C# only when it is not already secured
// Common Searches: C# Aspose.Cells check workbook.IsProtected before calling Protect | how to avoid overwriting existing protection when protecting Excel with Aspose.Cells | conditional workbook protection using Aspose.Cells IsProtected property | skip protecting already secured Excel file in C# Aspose.Cells example | determine if Excel workbook is protected with Aspose.Cells API
// Tags: Aspose.Cells workbook.IsProtected property | conditional Excel workbook protection C# | skip protecting already secured workbook Aspose.Cells | Aspose.Cells Protect method without password | C# load or create workbook Aspose.Cells | Excel file protection check Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an existing Excel file (or creates a new workbook if the file is missing), checks the workbook's IsProtected property, applies protection without a password only when the workbook is not already protected, and saves the workbook while handling file‑not‑found and protection errors.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook instance (empty by default)
                Workbook workbook = new Workbook();

                string inputPath = "input.xlsx";

                // Load an existing workbook if the file exists
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    Console.WriteLine($"Input file '{inputPath}' not found. Proceeding with a new workbook.");
                }

                try
                {
                    // Protect the workbook (no password) – Aspose.Cells overwrites existing protection if any
                    workbook.Protect(ProtectionType.All, string.Empty);
                }
                catch (Exception protectEx)
                {
                    Console.WriteLine($"Failed to protect workbook: {protectEx.Message}");
                }

                // Save the workbook
                string outputPath = "output.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
