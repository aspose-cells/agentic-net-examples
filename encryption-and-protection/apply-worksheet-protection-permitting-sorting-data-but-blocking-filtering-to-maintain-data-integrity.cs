// Title: How to protect an Excel worksheet while allowing sorting and disabling filtering with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# protect worksheet allow sorting but disable filter | set worksheet protection password and enable sorting only using Aspose.Cells | C# code to enable sorting on protected sheet while blocking auto‑filter in Excel | how to configure worksheet protection options in Aspose.Cells .NET | prevent users from applying filters on a protected worksheet with Aspose.Cells
// Tags: worksheet protection allow sorting Aspose.Cells | disable filtering on protected Excel sheet C# | set worksheet password Aspose.Cells .NET | configure protection options Excel workbook Aspose.Cells | protect worksheet sorting enabled filtering disabled

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing workbook, accesses the first worksheet, configures its Protection object to permit sorting while blocking filtering, assigns a password, and saves the modified workbook.
class WorksheetProtectionExample
{
    static void Main()
    {
        try
        {
            // Define input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the protection object (read‑only property)
            var protection = worksheet.Protection;

            // Permit sorting of data
            protection.AllowSorting = true;

            // Block filtering to maintain data integrity
            protection.AllowFiltering = false;

            // Set a password for the protection (use the Password property)
            protection.Password = "securePassword";

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
