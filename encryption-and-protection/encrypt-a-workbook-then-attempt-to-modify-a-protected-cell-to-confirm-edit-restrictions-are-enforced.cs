// Title: Encrypt an Excel workbook with a password and verify that worksheet protection blocks cell edits using Aspose.Cells for .NET
// Common Searches: C# encrypt Excel file with password using Aspose.Cells while keeping sheet protection active | Aspose.Cells protect worksheet and prevent editing of locked cells after saving encrypted workbook | Load a password‑protected XLSX with Aspose.Cells and test if cell lock is still enforced | How to catch CellsException when trying to edit a protected cell in an encrypted workbook with Aspose.Cells
// Tags: encrypt workbook with password Aspose.Cells C# | worksheet protection enforce cell lock Aspose.Cells | load password‑protected XLSX using LoadOptions Aspose.Cells | handle CellsException for protected cell edit Aspose.Cells | save encrypted Excel file Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, protects the first worksheet, attempts to edit a locked cell (capturing the expected CellsException), sets a workbook password to encrypt the file, saves it, reloads the encrypted workbook with LoadOptions, and tries to edit the same cell again to confirm that protection remains enforced after decryption.
class WorkbookProtectionDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and access the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "DataSheet";

            // Put an initial value into cell A1 (default cells are locked)
            Cell targetCell = sheet.Cells["A1"];
            targetCell.PutValue("Original Value");

            // Protect the worksheet to enforce cell lock (oldPassword not required)
            sheet.Protect(ProtectionType.All, "sheetPassword", null);

            // Attempt to modify the protected cell A1
            try
            {
                targetCell.PutValue("Attempted Edit");
                Console.WriteLine("Cell edited successfully (unexpected).");
            }
            catch (CellsException)
            {
                Console.WriteLine("Edit prevented: cell is protected.");
            }

            // Encrypt the workbook with a password
            workbook.Settings.Password = "workbookPassword";

            // Save the encrypted workbook to disk
            string filePath = "EncryptedProtectedWorkbook.xlsx";
            try
            {
                workbook.Save(filePath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved and encrypted at: {filePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving workbook: {ex.Message}");
                return;
            }

            // Verify the file exists before loading
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            // Load the encrypted workbook using the password
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = "workbookPassword"
            };

            Workbook loadedWorkbook;
            try
            {
                loadedWorkbook = new Workbook(filePath, loadOptions);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading workbook: {ex.Message}");
                return;
            }

            Worksheet loadedSheet = loadedWorkbook.Worksheets["DataSheet"];
            Cell loadedCell = loadedSheet.Cells["A1"];

            // Attempt to modify the protected cell after loading
            try
            {
                loadedCell.PutValue("Second Attempted Edit");
                Console.WriteLine("Cell edited after load (unexpected).");
            }
            catch (CellsException)
            {
                Console.WriteLine("Edit prevented after load: cell is still protected.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
