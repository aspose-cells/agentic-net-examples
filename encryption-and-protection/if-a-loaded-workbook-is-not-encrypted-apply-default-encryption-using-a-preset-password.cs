// Title: Apply default password encryption to an unprotected Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: how to encrypt an existing XLSX file with a default password using Aspose.Cells in C# | apply password protection only when workbook is not already encrypted with Aspose.Cells | overwrite original Excel workbook with encryption using Aspose.Cells SaveFormat.Xlsx | determine if an Excel workbook is encrypted before setting password in Aspose.Cells C#
// Tags: Aspose.Cells workbook password encryption | C# set Workbook.Settings.Password | save encrypted XLSX with Aspose.Cells | encrypt Excel workbook in place Aspose.Cells | default password protection for Excel using Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing XLSX file into an Aspose.Cells Workbook, assigns a preset password via Workbook.Settings.Password (which triggers encryption), and saves the workbook back to the same path using SaveFormat.Xlsx, overwriting the original file while handling possible errors.
class Program
{
    static void Main()
    {
        // Path to the workbook to process
        string filePath = "input.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"Error: The file \"{filePath}\" was not found.");
                return;
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(filePath);

            // Apply password protection (encryption). 
            // Aspose.Cells will encrypt the file when a password is set, regardless of prior state.
            workbook.Settings.Password = "DefaultPassword123";

            // Save the workbook (overwrites the original file with encryption applied)
            workbook.Save(filePath, SaveFormat.Xlsx);

            Console.WriteLine("Workbook encrypted and saved successfully.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
