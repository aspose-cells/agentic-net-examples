// Title: C# example to compare XLSX file size before and after Workbook.RemoveUnusedStyles using Aspose.Cells
// Common Searches: how to get Excel file size in bytes with Aspose.Cells before removing unused styles | C# Aspose.Cells remove unused styles and determine resulting workbook size | measure size reduction after calling Workbook.RemoveUnusedStyles in .NET | Aspose.Cells get workbook size without saving to disk using memory stream | detect password protected workbook when cleaning up styles with Aspose.Cells
// Tags: remove unused styles workbook optimization Aspose.Cells | calculate XLSX byte size C# Aspose.Cells | save workbook to memory stream Aspose.Cells | detect password protected workbook Aspose.Cells | compare original and optimized workbook size Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an XLSX workbook, records its on‑disk byte size, calls Workbook.RemoveUnusedStyles() to delete styles not referenced by any cell, saves the workbook to a MemoryStream in Xlsx format, and then outputs the new byte count and the reduction amount, while also handling password‑protected files via CellsException.
class Program
{
    static void Main()
    {
        // Path to the source workbook
        string inputPath = "input.xlsx";

        // Verify that the input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook (no password supplied)
            Workbook workbook = new Workbook(inputPath);

            // Measure original file size on disk
            long originalSize = new FileInfo(inputPath).Length;
            Console.WriteLine($"Original file size: {originalSize} bytes");

            // Remove styles that are not used by any cell
            workbook.RemoveUnusedStyles();

            // Save the modified workbook to a memory stream to obtain its new size
            using (MemoryStream ms = new MemoryStream())
            {
                workbook.Save(ms, SaveFormat.Xlsx);
                long newSize = ms.Length;
                Console.WriteLine($"File size after removing unused styles: {newSize} bytes");
                Console.WriteLine($"Size reduction: {originalSize - newSize} bytes");
            }
        }
        catch (CellsException ex)
        {
            // Detect password protection via exception message
            if (ex.Message != null && ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                Console.WriteLine("The workbook is password protected. Please provide a password.");
            }
            else
            {
                Console.WriteLine($"CellsException: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
