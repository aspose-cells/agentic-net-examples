// Title: Replace the modify‑only password of an existing XLSX workbook using Aspose.Cells for .NET without changing its data
// AI Prompts: Load an existing .xlsx file with Aspose.Cells, apply a new modify‑only password via Workbook.Protect, and save the workbook unchanged. | Use C# and Aspose.Cells to change the protection password that restricts modifications on an Excel file while preserving all worksheets. | Programmatically update the modify password of a protected workbook without altering its content using the Aspose.Cells API.
// Common Searches: Aspose.Cells C# change modify password of existing Excel file | How to set a new modify‑only password on an XLSX workbook using Aspose.Cells | Replace workbook protection password without modifying data in .NET | Update Excel file modify password programmatically with Aspose.Cells | Change protection password type All for existing workbook Aspose.Cells
// Tags: Aspose.Cells workbook.Protect modify password | replace modify password XLSX C# | set protection type all Aspose.Cells | preserve worksheet data while protecting workbook | load workbook with LoadOptions Aspose.Cells | save workbook unchanged Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing XLSX workbook using Aspose.Cells, assigns a new modify‑only password with Workbook.Protect(ProtectionType.All, newPassword), and saves the file to a new location while leaving all worksheet content untouched.
class Program
{
    static void Main()
    {
        // Path to the existing workbook
        string inputPath = "input.xlsx";

        // New password that will be required to modify the workbook
        string newModifyPassword = "newPassword";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook (no opening password is assumed)
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
            // If the workbook has an opening password, set it here:
            // loadOptions.Password = "openingPassword";

            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Set the password required to modify the workbook.
            // Use Protect with ProtectionType.All to require a modify password.
            workbook.Protect(ProtectionType.All, newModifyPassword);

            // Save the workbook without altering its content
            string outputPath = "output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
