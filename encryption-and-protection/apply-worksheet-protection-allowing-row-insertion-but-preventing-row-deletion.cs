// Title: Protect an Excel worksheet with Aspose.Cells for .NET allowing row insertion but blocking row deletion
// Common Searches: C# Aspose.Cells protect worksheet and enable only row insertion | How to set Excel sheet protection to allow inserting rows but prevent deleting rows using Aspose.Cells | Aspose.Cells .NET example for custom sheet protection permissions for rows | Apply password protection to a worksheet while allowing row inserts in Aspose.Cells | Configure worksheet protection options in Aspose.Cells to restrict row deletion
// Tags: worksheet protection insert rows Aspose.Cells | disable row deletion Aspose.Cells | Aspose.Cells set sheet protection options | Excel worksheet password protection .NET | custom sheet protection permissions Aspose.Cells | Aspose.Cells protect worksheet rows

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing Excel file, accesses the first worksheet, configures its protection to permit row insertion while forbidding row deletion, applies password protection, and saves the protected workbook using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";
            const string password = "myPassword";

            // Ensure the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Configure protection options
            sheet.Protection.AllowInsertingRow = true;
            sheet.Protection.AllowDeletingRow = false;

            // Apply protection (oldPassword is empty because the sheet is not previously protected)
            sheet.Protect(ProtectionType.All, password, string.Empty);

            // Save the protected workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
