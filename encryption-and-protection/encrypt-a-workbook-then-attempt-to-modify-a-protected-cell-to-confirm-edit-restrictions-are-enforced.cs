// Title: How to encrypt an Excel workbook, protect a worksheet, and verify locked‑cell edit restrictions using Aspose.Cells for .NET
// AI Prompts: Create a new workbook, apply a worksheet password, encrypt the file with a workbook password, then attempt to write to a locked cell and capture the resulting CellsException. | Open a password‑protected Excel file via LoadOptions, modify only unlocked cells, and save the workbook while preserving its encryption.
// Common Searches: asp.net encrypt excel file and protect sheet with Aspose.Cells | prevent editing of locked cells after workbook password protection using Aspose.Cells | catch CellsException when modifying a protected cell in an encrypted workbook Aspose.Cells | load encrypted Excel workbook with password and save changes without removing protection Aspose.Cells
// Tags: apply workbook password encryption Aspose.Cells | set worksheet protection password Aspose.Cells | prevent edit of locked cells Aspose.Cells | load password‑protected Excel with LoadOptions Aspose.Cells | save modified workbook retaining encryption Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, protects the first worksheet with a password, encrypts the file using a workbook password, reloads it with LoadOptions, attempts to modify a locked cell (capturing the expected CellsException), successfully updates an unlocked cell, and saves the changes while keeping the workbook encrypted.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Sheet1";

            // Populate cells: A1 will be editable, B1 will be protected
            sheet.Cells["A1"].PutValue("Editable");
            sheet.Cells["B1"].PutValue("Protected");

            // Protect the entire sheet with a password (cells are locked by default)
            // The third parameter is the old password; not needed here, so pass null
            sheet.Protect(ProtectionType.All, "sheetpwd", null);

            // Encrypt the workbook with a password
            workbook.Settings.Password = "workbookpwd";

            // Save the encrypted workbook
            string filePath = "EncryptedWorkbook.xlsx";
            workbook.Save(filePath);
            Console.WriteLine($"Workbook saved to {filePath}");

            // Ensure the file exists before attempting to load it
            if (!File.Exists(filePath))
                throw new FileNotFoundException("The workbook file was not found.", filePath);

            // Load the workbook using the encryption password
            LoadOptions loadOptions = new LoadOptions
            {
                Password = "workbookpwd"
            };
            Workbook loadedWorkbook = new Workbook(filePath, loadOptions);
            Worksheet loadedSheet = loadedWorkbook.Worksheets[0];

            // Attempt to modify the protected cell B1
            try
            {
                loadedSheet.Cells["B1"].PutValue("Attempted Change");
                Console.WriteLine("Protected cell modified successfully (unexpected).");
            }
            catch (CellsException ex)
            {
                // Expected exception because the cell is locked
                Console.WriteLine($"Modification failed as expected: {ex.Message}");
            }

            // Modify the editable cell A1 to demonstrate allowed changes
            loadedSheet.Cells["A1"].PutValue("Changed");
            Console.WriteLine("Editable cell changed successfully.");

            // Save the workbook after modifications (optional)
            string modifiedPath = "ModifiedWorkbook.xlsx";
            loadedWorkbook.Save(modifiedPath);
            Console.WriteLine($"Modified workbook saved to {modifiedPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
