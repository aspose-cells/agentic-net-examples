// Title: Protect a single worksheet with a password and save the workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells protect only one sheet with a password and keep other sheets editable | How to set worksheet protection password in Aspose.Cells .NET without encrypting the whole file | Saving an Excel workbook after applying worksheet.Protect in Aspose.Cells C# | Aspose.Cells protect sheet using ProtectionType.All example in C# | Apply password to specific worksheet in existing .xlsx using Aspose.Cells library
// Tags: worksheet.Protect password Aspose.Cells C# | protect single Excel sheet .NET | Aspose.Cells worksheet protection example | save workbook after sheet protection Aspose.Cells | Excel file per‑sheet password protection C#

using Aspose.Cells;
using System;
using System.IO;

// The code verifies the input.xlsx file, loads it with Aspose.Cells, accesses the "Sheet1" worksheet, applies password protection using worksheet.Protect with ProtectionType.All and the password "myPassword", and then saves the modified workbook as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Ensure the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the worksheet you want to protect
            Worksheet worksheet = workbook.Worksheets["Sheet1"]; // or use workbook.Worksheets[0]

            // Apply password protection (oldPassword is empty for new protection)
            worksheet.Protect(ProtectionType.All, "", "myPassword");

            // Save the workbook with the protected worksheet
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
