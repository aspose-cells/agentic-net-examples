// Title: Change worksheet protection password while retaining existing protection settings in an Excel file using Aspose.Cells for .NET
// AI Prompts: Generate C# code that opens an existing .xlsx workbook with Aspose.Cells, updates the worksheet's protection password, and saves the file without altering any other protection flags. | Show how to modify only the password of a protected worksheet in Aspose.Cells while preserving its current allowed actions such as editing objects or formatting cells. | Provide a step‑by‑step example that loads a workbook, sets a new worksheet protection password, and writes the result to a new file using Aspose.Cells for .NET.
// Common Searches: aspnet change worksheet protection password without resetting allowed actions Aspose.Cells | c# update Excel worksheet password keep existing protection options | how to preserve worksheet protection settings when changing password using Aspose.Cells | Aspose.Cells set new worksheet password while keeping protection flags
// Tags: worksheet protection password change Aspose.Cells | retain existing worksheet protection flags .NET | Excel worksheet password modification Aspose.Cells | Aspose.Cells preserve protection settings on password update | C# set new worksheet protection password

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an existing XLSX workbook, assigns a new password to the first worksheet's protection while keeping all other protection options unchanged, and saves the result to a new file using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";
            const string newPassword = "NewSecurePassword123";

            try
            {
                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Set the worksheet password while preserving existing protection settings
                sheet.Protection.Password = newPassword;

                // Save the updated workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
