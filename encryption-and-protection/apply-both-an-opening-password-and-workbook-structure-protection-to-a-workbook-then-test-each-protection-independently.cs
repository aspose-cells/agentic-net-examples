// Title: Apply an opening password and protect workbook structure in an Excel file with Aspose.Cells for .NET, then verify each protection
// Common Searches: how to set an opening password on an Excel workbook using Aspose.Cells in C# | protect workbook structure with Aspose.Cells to block adding new worksheets | load a password‑protected XLSX file with LoadOptions in Aspose.Cells .NET | unprotect workbook structure programmatically after Aspose.Cells protection | test that opening password enforcement works when opening a protected workbook with Aspose.Cells
// Tags: Aspose.Cells opening password protection | Aspose.Cells workbook structure protection | LoadOptions password parameter Aspose.Cells | Unprotect workbook structure Aspose.Cells | prevent worksheet addition after structure protection Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new workbook, assigns an opening password via Settings.Password, applies structure protection with Protect, saves the file, then attempts to open it without a password to confirm the opening protection. It reloads the workbook using LoadOptions with the correct password, tries to add a worksheet to verify that structure protection blocks the operation, unprotects the structure with Unprotect, adds a new worksheet successfully, and finally saves the modified workbook.
class WorkbookProtectionDemo
{
    static void Main()
    {
        try
        {
            // 1. Create a new workbook and add some sample data
            Workbook wb = new Workbook();
            Worksheet sheet = wb.Worksheets[0];
            sheet.Name = "DataSheet";
            sheet.Cells["A1"].PutValue("Hello");
            sheet.Cells["B1"].PutValue("World");

            // 2. Apply opening password and protect workbook structure
            wb.Settings.Password = "open123"; // opening password
            wb.Protect(ProtectionType.Structure, "struct123"); // structure protection

            // 3. Save the protected workbook
            string protectedPath = "ProtectedWorkbook.xlsx";
            wb.Save(protectedPath);
            Console.WriteLine($"Protected workbook saved to {protectedPath}");

            // -------------------------------------------------
            // 4. Test opening password protection (load without password)
            try
            {
                if (!File.Exists(protectedPath))
                    throw new FileNotFoundException("Protected workbook file not found.", protectedPath);

                // Attempt to load without providing a password
                Workbook testLoad = new Workbook(protectedPath);
                Console.WriteLine("Workbook opened without password (unexpected).");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Opening password protection works: " + ex.Message);
            }

            // -------------------------------------------------
            // 5. Load workbook with opening password
            LoadOptions loadOpts = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = "open123"
            };

            Workbook loadedWb;
            try
            {
                if (!File.Exists(protectedPath))
                    throw new FileNotFoundException("Protected workbook file not found.", protectedPath);

                loadedWb = new Workbook(protectedPath, loadOpts);
                Console.WriteLine("Workbook opened with correct opening password.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed to open workbook with password: " + ex.Message);
                return;
            }

            // 6. Test structure protection (attempt to add a sheet without unprotecting)
            try
            {
                loadedWb.Worksheets.Add("NewSheetWithoutUnprotect");
                Console.WriteLine("Added sheet without unprotecting structure (unexpected).");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Structure protection works (cannot add sheet): " + ex.Message);
            }

            // 7. Unprotect workbook structure using the correct password
            try
            {
                loadedWb.Unprotect("struct123");
                Console.WriteLine("Workbook structure unprotected.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed to unprotect workbook structure: " + ex.Message);
                return;
            }

            // 8. Now adding a new worksheet should succeed
            try
            {
                loadedWb.Worksheets.Add("NewSheetAfterUnprotect");
                Console.WriteLine("New worksheet added after unprotecting structure.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed to add worksheet after unprotecting: " + ex.Message);
            }

            // 9. Save the modified workbook
            string modifiedPath = "ModifiedWorkbook.xlsx";
            try
            {
                loadedWb.Save(modifiedPath);
                Console.WriteLine($"Modified workbook saved to {modifiedPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed to save modified workbook: " + ex.Message);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("An unexpected error occurred: " + ex.Message);
        }
    }
}
