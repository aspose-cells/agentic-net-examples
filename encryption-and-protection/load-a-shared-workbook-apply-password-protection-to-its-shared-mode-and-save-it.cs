// Title: Load an existing shared Excel workbook, enable shared mode, set an open password, and save it with Aspose.Cells for .NET
// Common Searches: how to set an open password on a shared Excel workbook using Aspose.Cells C# | Aspose.Cells enable shared mode and protect workbook with password in .NET | C# code to load a shared workbook, apply password protection, and save it with Aspose.Cells | protect shared Excel file with password while keeping shared mode using Aspose.Cells | Aspose.Cells workbook.Settings.Password example for shared workbooks
// Tags: Aspose.Cells set workbook open password | shared workbook protection Aspose.Cells | enable shared mode workbook.Settings.Shared | save password-protected Excel file Aspose.Cells | load existing workbook Aspose.Cells C#

using System;
using System.IO;
using Aspose.Cells;

// The example checks for the presence of 'SharedWorkbook.xlsx', loads it into an Aspose.Cells Workbook, enables shared mode via Settings.Shared, assigns an open password using Settings.Password, and saves the protected file as 'SharedWorkbook_Protected.xlsx', handling any exceptions that may occur.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "SharedWorkbook.xlsx";
            const string outputPath = "SharedWorkbook_Protected.xlsx";
            const string password = "YourPassword";

            // Verify that the source workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Enable shared mode
            workbook.Settings.Shared = true;

            // Apply password protection (open password) to the workbook
            workbook.Settings.Password = password;

            // Save the protected workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
