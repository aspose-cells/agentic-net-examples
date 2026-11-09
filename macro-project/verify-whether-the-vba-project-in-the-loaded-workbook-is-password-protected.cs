// Title: Check if an Excel workbook’s VBA project is password‑protected with Aspose.Cells for .NET (C#)
// Common Searches: how to determine if a VBA project in an .xlsx file is password protected using Aspose.Cells C# | Aspose.Cells C# read IsPasswordProtected property of VbaProject | detect macro project password protection in Excel with Aspose.Cells .NET | C# code to verify VBA project protection status in a workbook | using reflection to access VbaProject.IsPasswordProtected in older Aspose.Cells versions
// Tags: Aspose.Cells VbaProject password check | C# detect VBA project protection | Workbook.VbaProject.IsPasswordProtected reflection | Excel macro password detection Aspose.Cells | fallback for older Aspose.Cells versions

using Aspose.Cells;
using System;
using System.IO;

// The example loads an Excel workbook with Aspose.Cells, checks whether a VBA project exists, uses reflection to read the IsPasswordProtected property when available, outputs the protection status, and optionally saves the workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            bool isPasswordProtected = false;

            // Check if a VBA project exists
            if (workbook.VbaProject != null)
            {
                // Use reflection to safely access IsPasswordProtected (may not exist in older versions)
                var propInfo = workbook.VbaProject.GetType().GetProperty("IsPasswordProtected");
                if (propInfo != null && propInfo.PropertyType == typeof(bool))
                {
                    isPasswordProtected = (bool)propInfo.GetValue(workbook.VbaProject);
                }
            }

            // Output the result
            Console.WriteLine("VBA project password protected: " + isPasswordProtected);

            // Save the workbook (optional)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
