// Title: Protect a specific worksheet with a password and save the Excel workbook using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells protect a single worksheet with a password and save the file | How to apply worksheet protection by name using Aspose.Cells .NET | Saving an Excel workbook after setting sheet protection in Aspose.Cells C# | Load existing .xlsx, protect Sheet1 with password, and export using Aspose.Cells | Aspose.Cells protect sheet without affecting other worksheets C# example
// Tags: Worksheet.Protect API Aspose.Cells | Excel worksheet password protection C# | Saving workbook after sheet protection Aspose.Cells | Load .xlsx and protect specific sheet .NET | ProtectionType.All usage Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing .xlsx file into an Aspose.Cells Workbook, retrieves a worksheet by its name, applies password protection to the entire sheet using the Protect method with ProtectionType.All, and then saves the modified workbook to a new file, handling any errors that may occur.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";
            const string sheetName = "Sheet1";
            const string password = "myPassword";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the worksheet to protect (by name or index)
            Worksheet worksheet = workbook.Worksheets[sheetName];
            if (worksheet == null)
            {
                Console.WriteLine($"Worksheet \"{sheetName}\" not found.");
                return;
            }

            // Apply password protection to the worksheet.
            // The overload requires the old password; pass null when there is none.
            worksheet.Protect(ProtectionType.All, password, null);

            // Save the workbook with the protected worksheet
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
