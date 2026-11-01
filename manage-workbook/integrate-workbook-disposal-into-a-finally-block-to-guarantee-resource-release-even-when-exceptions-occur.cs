// Title: How to guarantee Aspose.Cells Workbook disposal in a C# finally block for exception‑safe Excel processing
// Common Searches: C# dispose Aspose.Cells Workbook in finally block after saving Excel file | ensure Excel workbook resources are released with Aspose.Cells when an exception occurs | best practice for releasing Aspose.Cells Workbook resources in try‑catch‑finally | how to use finally to clean up Aspose.Cells Workbook object in .NET | prevent memory leaks with Aspose.Cells Workbook disposal in C#
// Tags: Aspose.Cells Workbook.Dispose usage | C# finally block for Excel resource cleanup | exception‑safe workbook handling with Aspose.Cells | memory leak prevention in Aspose.Cells | try‑catch‑finally pattern for Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads or creates an Excel workbook with Aspose.Cells, writes data, saves the file, handles specific and generic exceptions, and guarantees that the Workbook object is disposed inside a finally block to reliably release resources and avoid memory leaks.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        Workbook workbook = null;

        try
        {
            // Load existing workbook if it exists; otherwise create a new one
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
            }

            // Example operation: write a value to cell A1
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Hello, Aspose!");

            // Save the modified workbook
            workbook.Save(outputPath);
        }
        catch (CellsException ex)
        {
            // Handle Aspose.Cells specific errors (e.g., password protected files)
            Console.WriteLine($"Aspose.Cells error: {ex.Message}");
        }
        catch (Exception ex)
        {
            // Handle other unexpected errors
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
        finally
        {
            // Ensure the workbook is disposed
            if (workbook != null)
            {
                workbook.Dispose();
            }
        }
    }
}
