// Title: Read the address of a global named range on Sheet3 using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells get address of a global named range defined on Sheet3 | How to retrieve the RefersTo formula of a workbook global name with Aspose.Cells .NET | Access workbook.Worksheets.Names collection to read a named range address in C# | Aspose.Cells read global named range location without opening Excel | Get cell range address from a global name in an Excel file using Aspose.Cells for .NET
// Tags: global named range address Aspose.Cells | Workbook.Worksheets.Names collection C# | RefersTo property read formula Aspose.Cells | read named range location .NET | Aspose.Cells retrieve named range address

using Aspose.Cells;
using System;
using System.IO;

// The example loads an Excel workbook with Aspose.Cells, accesses a global named range via the workbook.Worksheets.Names collection, extracts its RefersTo formula (optionally stripping the leading '='), prints the cleaned address, and saves the workbook if needed.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";
        const string rangeName = "MyGlobalRange";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Retrieve a global named range from the workbook's Names collection
            Name globalName = workbook.Worksheets.Names[rangeName];

            if (globalName != null)
            {
                // RefersTo returns a formula string like "=Sheet3!$A$1:$B$5"
                string address = globalName.RefersTo;

                // Remove the leading '=' for a cleaner display (optional)
                if (address.StartsWith("="))
                    address = address.Substring(1);

                Console.WriteLine($"Global named range address: {address}");
            }
            else
            {
                Console.WriteLine($"The specified global named range \"{rangeName}\" was not found.");
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook (optional – only if modifications were made)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
