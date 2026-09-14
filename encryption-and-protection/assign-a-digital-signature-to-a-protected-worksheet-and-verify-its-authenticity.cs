// Title: How to password‑protect an Excel worksheet, save as .xlsx, and programmatically verify the workbook using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that applies full protection to the first worksheet using a password, saves the workbook as an .xlsx file, and then reloads the file to confirm it opens without errors. | Show a .NET example that creates a workbook, renames the first sheet, protects it with a password, persists the workbook, checks file existence, and validates successful loading using Aspose.Cells.
// Common Searches: Aspose.Cells C# example to protect a worksheet with a password and verify the saved file | save password‑protected Excel workbook and reload it using Aspose.Cells for .NET | how to check if a protected Excel sheet can be opened after saving with Aspose.Cells | C# code to apply worksheet protection and test workbook integrity with Aspose.Cells
// Tags: worksheet protection with password Aspose.Cells C# | save protected workbook as xlsx Aspose.Cells | load and validate Excel file Aspose.Cells .NET | verify worksheet protection after save Aspose.Cells | C# Aspose.Cells workbook integrity check

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a new workbook, renames the first worksheet, applies full protection using a password, saves the workbook as an .xlsx file, confirms the file exists, and reloads it to ensure the protected workbook opens correctly, handling any exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "ProtectedSheet";

            // Protect the worksheet with a password
            sheet.Protect(ProtectionType.All);
            sheet.Protection.Password = "sheetPassword";

            // Save the workbook
            string outputFile = "ProtectedWorkbook.xlsx";
            workbook.Save(outputFile, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully: {outputFile}");

            // Verify that the file was created
            if (File.Exists(outputFile))
            {
                // Load the saved workbook to ensure it can be opened
                Workbook loadedWorkbook = new Workbook(outputFile);
                Console.WriteLine("Workbook loaded successfully.");
            }
            else
            {
                Console.WriteLine($"Failed to locate the saved workbook: {outputFile}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
