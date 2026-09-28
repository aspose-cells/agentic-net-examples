// Title: Check if an .xlsm workbook's VBA project is digitally signed using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsm file with Aspose.Cells and uses Workbook.VbaProject.IsSigned to report whether the VBA macro has a digital signature. | Create a .NET console example that verifies the presence of a VBA project signature in an Excel workbook and gracefully handles missing files or errors.
// Common Searches: asp.net c# aspose.cells determine if VBA macro is signed in xlsm | how to check VBA project digital signature with Aspose.Cells | Workbook.VbaProject.IsSigned example c# console | detect unsigned VBA project in Excel using Aspose.Cells .NET | validate VBA macro signature Aspose.Cells C# snippet
// Tags: check VBA project signature Aspose.Cells | Workbook.VbaProject.IsSigned usage | detect unsigned VBA macro .xlsm | load Excel workbook and inspect VBA signature .NET | Aspose.Cells VBA digital signature validation

using Aspose.Cells;
using System;
using System.IO;

// The example loads an .xlsm workbook with Aspose.Cells, verifies that the file exists, checks if the workbook contains a VBA project and whether it is digitally signed via the IsSigned property, and outputs the result to the console while handling exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsm";

        try
        {
            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: File not found - {inputPath}");
                return;
            }

            // Load the workbook that contains a VBA project
            Workbook workbook = new Workbook(inputPath);

            // Check whether the VBA project is signed (digital signature present)
            bool isSigned = workbook.VbaProject != null && workbook.VbaProject.IsSigned;

            // Display the result
            Console.WriteLine("VBA project signed: " + isSigned);
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
