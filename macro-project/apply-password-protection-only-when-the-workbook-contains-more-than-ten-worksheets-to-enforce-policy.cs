// Title: Apply password protection to an Excel workbook with Aspose.Cells in C# only when it contains more than ten worksheets
// AI Prompts: Write a C# program using Aspose.Cells that opens an .xlsx file, checks if workbook.Worksheets.Count exceeds 10, and calls workbook.Protect(ProtectionType.All, password) before saving. | Update an existing Aspose.Cells console application to add a conditional block that applies a strong password to the workbook only when the sheet count is greater than ten. | Create a .NET snippet that validates the input path, ensures the output directory exists, and conditionally protects the workbook based on the number of worksheets.
// Common Searches: aspocells protect workbook when sheet count greater than ten c# | conditional workbook password protection aspnet aspocells | how to apply Aspose.Cells workbook.Protect only for large Excel files | c# check number of worksheets before encrypting Excel with Aspose.Cells
// Tags: Aspose.Cells conditional workbook protection | C# workbook.Protect based on worksheet count | protect Excel file with password using Aspose.Cells | check worksheet count before applying protection | save protected workbook to new file Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// // Loads an existing .xlsx file, verifies the worksheet count, applies password protection with ProtectionType.All only when more than ten sheets are present, and saves the protected workbook to a new location.
class WorkbookProtection
{
    static void Main()
    {
        // Define input and output file paths
        string inputPath = @"C:\Input\Sample.xlsx";
        string outputPath = @"C:\Output\Sample_Protected.xlsx";

        // Verify that the input workbook exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Apply password protection if the workbook has more than ten worksheets
            if (workbook.Worksheets.Count > 10)
            {
                string password = "StrongPassword123";
                // Protect the entire workbook with the specified password
                workbook.Protect(ProtectionType.All, password);
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the protected workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
