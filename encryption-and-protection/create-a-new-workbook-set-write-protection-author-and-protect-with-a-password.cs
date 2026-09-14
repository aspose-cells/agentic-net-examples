// Title: Create an Excel workbook, assign author metadata, and protect it with a password using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that uses Aspose.Cells to generate a new workbook, set the Author property, and apply full write protection with a password. | Show how to protect the workbook structure and windows while adding author metadata in Aspose.Cells. | Demonstrate saving the protected workbook to a specified path and automatically creating the output directory if it does not exist.
// Common Searches: Aspose.Cells C# set workbook author and password protect the file | How to apply write protection to an Excel workbook using Aspose.Cells in .NET | Create a new Excel workbook and protect it with a password with Aspose.Cells for C# | Save Aspose.Cells workbook to a folder that may not exist | Protect workbook structure and windows with password using Aspose.Cells API
// Tags: Aspose.Cells workbook password protection | Aspose.Cells set author property | C# create protected Excel file with Aspose.Cells | Aspose.Cells auto-create output folder on save | protect workbook structure and windows Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// // Creates a new Workbook, assigns the Author metadata, applies write protection for all elements with a password, ensures the target directory exists, and saves the file as ProtectedWorkbook.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Set workbook author (optional metadata)
            workbook.Settings.Author = "Jane Smith";

            // Apply write protection with password; protects structure and windows
            workbook.Protect(ProtectionType.All, "MySecurePassword");

            // Define output file path
            string outputPath = "ProtectedWorkbook.xlsx";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the protected workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
