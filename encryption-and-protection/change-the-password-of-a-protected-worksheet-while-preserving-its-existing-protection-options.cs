// Title: Change the password of a protected worksheet while preserving its protection settings with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# change worksheet protection password without resetting options | update Excel worksheet password programmatically while keeping existing protection settings | how to modify worksheet protection password using Aspose.Cells .NET API | preserve worksheet protection options when changing password in C# | set new password for protected sheet in Aspose.Cells without losing protection flags
// Tags: worksheet protection password change Aspose.Cells | preserve worksheet protection settings C# | update Excel sheet password Aspose.Cells .NET | modify worksheet protection options programmatically | Aspose.Cells worksheet protection API

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Loads an existing workbook, updates the password of a protected worksheet while retaining all other protection options, and saves the modified file.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";
            const string newPassword = "NewSecurePassword123";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet (index 0)
                Worksheet sheet = workbook.Worksheets[0];

                // Change the worksheet protection password while preserving other protection settings
                // Setting the Password property updates only the password
                sheet.Protection.Password = newPassword;

                // Save the workbook with the updated protection settings
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
