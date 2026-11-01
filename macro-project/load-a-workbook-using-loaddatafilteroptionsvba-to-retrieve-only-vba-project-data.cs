// Title: How to load an Excel workbook and access its VBA project with Aspose.Cells in C#
// Common Searches: C# Aspose.Cells load workbook and read VBA project without opening Excel | retrieve macro code from XLSX using Aspose.Cells LoadOptions in .NET | access VbaProject object after loading Excel file with Aspose.Cells C# example | read VBA project from Excel file programmatically using Aspose.Cells library | save workbook after extracting VBA project using Aspose.Cells in C#
// Tags: Aspose.Cells load workbook VBA project C# | VbaProject access using LoadOptions | extract macro data from XLSX Aspose.Cells | programmatic VBA project retrieval .NET | save workbook after VBA extraction Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Vba;
using System;
using System.IO;

// The example demonstrates loading an XLSX workbook with Aspose.Cells LoadOptions in C#, accessing the workbook's VbaProject to work with macro code, and optionally saving the workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook (VBA project will be loaded automatically if present)
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Access the VBA project (if needed)
            VbaProject vbaProject = workbook.VbaProject;

            // Save the workbook (optional)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
