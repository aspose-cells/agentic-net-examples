// Title: How to set an opening password and worksheet structure protection on an Excel file with Aspose.Cells for .NET and verify each protection
// AI Prompts: Create C# code that applies a file‑open password using Workbook.Settings.Password and protects the worksheet structure with Worksheet.Protect, then saves the workbook. | Write C# logic to load the password‑protected workbook with LoadOptions, attempt opening without a password, and test unprotecting the worksheet with both incorrect and correct passwords, handling CellsException.
// Common Searches: aspnet set file open password for Excel using Aspose.Cells | protect worksheet structure with password Aspose.Cells C# example | load password protected xlsx with Aspose.Cells LoadOptions .NET | catch CellsException when opening Excel without password Aspose.Cells | unprotect worksheet with wrong password Aspose.Cells handling exception
// Tags: Workbook.Settings.Password protection Aspose.Cells | Worksheet.Protect structure password C# | LoadOptions password opening Excel Aspose.Cells | CellsException handling for protection errors | save protected workbook Xlsx Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // // Demonstrates creating a workbook, applying an opening password via Workbook.Settings.Password, protecting the worksheet structure with Worksheet.Protect, saving the file, attempting to open it without a password (expecting a CellsException), loading it with the correct password, testing unprotect with an incorrect password (catching the exception), then unprotecting with the correct password and saving an unprotected copy.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and add some data
                Workbook workbook = new Workbook();
                workbook.Worksheets[0].Cells["A1"].PutValue("Sample Data");

                // ---------- Apply opening password ----------
                workbook.Settings.Password = "OpenPwd123"; // password required to open the file

                // ---------- Apply worksheet protection (acts as structure protection for demo) ----------
                Worksheet sheet = workbook.Worksheets[0];

                // Protect the worksheet with a password (oldPassword is not required for new protection)
                sheet.Protect(ProtectionType.All, "StructPwd456", string.Empty);

                // Save the protected workbook
                string filePath = "ProtectedWorkbook.xlsx";
                workbook.Save(filePath, SaveFormat.Xlsx);

                // ---------- Test opening password ----------
                try
                {
                    // Attempt to load without providing a password – should fail
                    if (File.Exists(filePath))
                    {
                        Workbook wbNoPwd = new Workbook(filePath);
                        Console.WriteLine("ERROR: Workbook opened without password (unexpected).");
                    }
                    else
                    {
                        Console.WriteLine("ERROR: File not found: " + filePath);
                    }
                }
                catch (CellsException ex)
                {
                    Console.WriteLine("Opening without password failed as expected: " + ex.Message);
                }

                // Load with correct opening password
                if (File.Exists(filePath))
                {
                    LoadOptions loadOpts = new LoadOptions(LoadFormat.Xlsx)
                    {
                        Password = "OpenPwd123"
                    };
                    Workbook wbWithPwd = new Workbook(filePath, loadOpts);
                    Console.WriteLine("Workbook opened successfully with correct opening password.");

                    // ---------- Test worksheet protection ----------
                    Worksheet loadedSheet = wbWithPwd.Worksheets[0];

                    // Attempt to unprotect with wrong password – should raise an exception
                    try
                    {
                        loadedSheet.Unprotect("WrongPwd");
                        Console.WriteLine("ERROR: Worksheet unprotected with wrong password (unexpected).");
                    }
                    catch (CellsException ex)
                    {
                        Console.WriteLine("Unprotecting with wrong password failed as expected: " + ex.Message);
                    }

                    // Unprotect with correct password
                    loadedSheet.Unprotect("StructPwd456");
                    Console.WriteLine("Worksheet protection successfully removed with correct password.");

                    // (Optional) Save the workbook after removing protection
                    string unprotectedPath = "UnprotectedWorkbook.xlsx";
                    wbWithPwd.Save(unprotectedPath, SaveFormat.Xlsx);
                    Console.WriteLine("Unprotected workbook saved to: " + unprotectedPath);
                }
                else
                {
                    Console.WriteLine("ERROR: File not found for loading with password: " + filePath);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An unexpected error occurred: " + ex.Message);
            }
        }
    }
}
