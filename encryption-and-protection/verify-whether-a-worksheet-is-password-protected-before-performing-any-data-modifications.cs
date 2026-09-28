// Title: Check if an Excel worksheet is password‑protected before modifying cells using Aspose.Cells for .NET
// AI Prompts: Generate C# code with Aspose.Cells that determines whether a worksheet is locked with a password and only writes to a cell when the sheet is unprotected. | Show how to catch protection‑related exceptions in Aspose.Cells and conditionally skip cell updates for a protected worksheet.
// Common Searches: C# Aspose.Cells how to know if a worksheet is password protected before editing | detect protected sheet in Excel using Aspose.Cells .NET | skip cell update when worksheet is locked with Aspose.Cells | handle protected worksheet exception Aspose.Cells C# example | check sheet protection status programmatically Aspose.Cells
// Tags: worksheet protection detection Aspose.Cells C# | conditional cell write based on sheet lock | Aspose.Cells protected worksheet handling | Excel sheet password check .NET | exception handling for locked worksheet Aspose

using System;
using System.IO;
using Aspose.Cells;

// The program loads an Excel workbook with Aspose.Cells, attempts to modify a cell inside a try‑catch block, interprets any exception as an indication that the worksheet is password‑protected, reports the outcome, and saves the workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook;
            try
            {
                workbook = new Workbook(inputPath);
            }
            catch (Exception loadEx)
            {
                Console.WriteLine($"Failed to load workbook: {loadEx.Message}");
                return;
            }

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Attempt to modify a cell; if the sheet is protected an exception will be thrown
            try
            {
                worksheet.Cells["A1"].PutValue("Modified value");
                Console.WriteLine("Cell A1 updated successfully.");
            }
            catch (Exception modifyEx)
            {
                Console.WriteLine($"Worksheet may be protected; no changes were made. Details: {modifyEx.Message}");
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
