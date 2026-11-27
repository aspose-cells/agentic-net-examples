// Title: Check if an Excel workbook's VBA project is digitally signed using Aspose.Cells for .NET (C#)
// Common Searches: how to determine if a VBA project is signed in an .xlsx using Aspose.Cells C# | Aspose.Cells C# detect VBA macro digital signature status | read VBA project IsSigned property from Excel workbook with Aspose.Cells | verify presence of VBA digital signature in Excel file using Aspose.Cells library | C# code to identify signed VBA project in a workbook via Aspose.Cells
// Tags: Aspose.Cells VBA IsSigned property | detect signed VBA project in .xlsx | C# check VBA macro signature status | Excel workbook VBA digital signature detection | Aspose.Cells workbook VBA project verification

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, confirms a VBA project exists, and uses the VbaProject.IsSigned property to report whether the VBA project is digitally signed, noting that Aspose.Cells does not provide a separate signature validation API.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Ensure the workbook contains a VBA project
            if (workbook.VbaProject != null)
            {
                // Check whether the VBA project is digitally signed
                if (workbook.VbaProject.IsSigned)
                {
                    // Aspose.Cells does not expose a direct ValidateSignature method for VBA projects.
                    // The presence of IsSigned indicates the project is signed.
                    Console.WriteLine("The VBA project is digitally signed.");
                }
                else
                {
                    Console.WriteLine("The VBA project is not digitally signed.");
                }
            }
            else
            {
                Console.WriteLine("No VBA project found in the workbook.");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors during processing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
