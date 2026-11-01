// Title: Set an opening password and protect workbook structure, then verify each protection with Aspose.Cells for .NET
// Common Searches: how to add an opening password to an Excel file using Aspose.Cells C# | protect workbook structure with a password in Aspose.Cells .NET | load a password protected workbook and catch IncorrectPasswordException with Aspose.Cells | check if workbook structure protection prevents adding worksheets in Aspose.Cells | unprotect workbook structure programmatically after opening password in Aspose.Cells
// Tags: set opening password Aspose.Cells | protect workbook structure Aspose.Cells | load password protected workbook Aspose.Cells | validate structure protection Aspose.Cells | unprotect workbook structure Aspose.Cells | handle IncorrectPasswordException Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, assigns an opening password via wb.Settings.Password, applies structure protection with wb.Protect(ProtectionType.Structure, ...), saves the file, attempts to open it without a password to trigger IncorrectPasswordException, reloads it with the correct password, tries to add a worksheet to confirm structure protection throws InvalidOperationException, then unprotects the structure using wb.Unprotect and successfully adds a new sheet before saving the final workbook.
class WorkbookProtectionDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add sample data
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            ws.Cells["A1"].PutValue("Sample Data");

            // ---------- Apply protections ----------
            // 1. Set opening password (required to open the file)
            wb.Settings.Password = "OpenPassword123";

            // 2. Protect workbook structure with a separate password
            //    This prevents adding, removing, renaming, or moving worksheets.
            wb.Protect(ProtectionType.Structure, "StructurePwd456");

            // Save the protected workbook
            string filePath = "ProtectedWorkbook.xlsx";
            wb.Save(filePath);
            Console.WriteLine($"Workbook saved to '{filePath}'.");

            // ---------- Test opening password ----------
            try
            {
                // Attempt to load without providing the opening password – should fail
                Workbook testOpenFail = new Workbook(filePath);
                Console.WriteLine("ERROR: Workbook opened without password (unexpected).");
            }
            catch (Exception ex) when (ex.GetType().Name == "IncorrectPasswordException")
            {
                Console.WriteLine("Opening password correctly enforced – cannot open without password.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected exception while opening without password: {ex.Message}");
            }

            // Load with the correct opening password
            LoadOptions loadOpts = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = "OpenPassword123"
            };

            if (!File.Exists(filePath))
                throw new FileNotFoundException($"The file '{filePath}' was not found.");

            Workbook wbOpened = new Workbook(filePath, loadOpts);
            Console.WriteLine("Workbook opened successfully with correct opening password.");

            // ---------- Test structure protection ----------
            try
            {
                // Attempt to add a new worksheet while structure protection is active – should fail
                wbOpened.Worksheets.Add("NewSheet");
                Console.WriteLine("ERROR: Worksheet added despite structure protection (unexpected).");
            }
            catch (InvalidOperationException)
            {
                Console.WriteLine("Structure protection correctly enforced – cannot add worksheet.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected exception while adding worksheet: {ex.Message}");
            }

            // To modify the workbook structure, unprotect it using the structure password
            wbOpened.Unprotect("StructurePwd456");

            // Now adding a worksheet should succeed
            wbOpened.Worksheets.Add("NewSheet");
            Console.WriteLine("Worksheet added after unprotecting structure.");

            // Save the final workbook (optional)
            string finalPath = "ProtectedWorkbook_Final.xlsx";
            wbOpened.Save(finalPath);
            Console.WriteLine($"Final workbook saved to '{finalPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
