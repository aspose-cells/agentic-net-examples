// Title: Check that a password‑protected .xlsx workbook cannot be opened by Excel 97‑2003 using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells verify encrypted .xlsx fails to load with LoadFormat.Excel97To2003 | how to detect incompatibility of modern .xlsx password encryption with Excel 97‑2003 in Aspose.Cells | Aspose.Cells LoadOptions Excel97To2003 throws exception for password‑protected workbook | test if older Excel versions can open a password‑encrypted .xlsx using Aspose.Cells .NET
// Tags: Aspose.Cells password encryption compatibility | LoadOptions Excel97To2003 encrypted file handling | C# verify encrypted .xlsx incompatibility with Excel 97‑2003 | exception for unsupported encrypted workbook in Aspose.Cells | test encrypted workbook opening failure in older Excel

using System;
using Aspose.Cells;

// The sample creates a workbook, adds data, applies a password using Aspose.Cells' default Xlsx encryption, saves it as an .xlsx file, then attempts to load the same file with LoadOptions set to Excel97To2003. Because the modern encryption is not supported by the older Excel format, an exception is thrown, which is caught and reported as the expected outcome. The test file is deleted afterward.
class WorkbookEncryptionVerification
{
    static void Main()
    {
        // Path for the encrypted workbook
        string encryptedFile = "EncryptedWorkbook.xlsx";

        // Create a new workbook and add some data
        Workbook wb = new Workbook();
        Worksheet sheet = wb.Worksheets[0];
        sheet.Cells["A1"].PutValue("Test Data");

        // Set a password for encryption (modern encryption, compatible with .xlsx)
        // This will encrypt the file using the default Xlsx encryption algorithm
        wb.Settings.Password = "SecretPassword";

        // Save the workbook as .xlsx (default format)
        wb.Save(encryptedFile, SaveFormat.Xlsx);

        // Attempt to load the encrypted workbook using LoadOptions that target the older Excel 97-2003 format
        // Older Excel versions (xls) cannot open files encrypted with the modern Xlsx encryption
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Excel97To2003);

        try
        {
            // This load should fail because the file is encrypted with a newer algorithm
            Workbook oldVersionWb = new Workbook(encryptedFile, loadOptions);
            Console.WriteLine("Unexpected: Older Excel version was able to open the encrypted workbook.");
        }
        catch (Exception ex)
        {
            // Expected path: an exception is thrown indicating incompatibility
            Console.WriteLine("As expected, older Excel versions cannot open the encrypted workbook.");
            Console.WriteLine("Exception message: " + ex.Message);
        }
        finally
        {
            // Clean up the test file
            if (System.IO.File.Exists(encryptedFile))
                System.IO.File.Delete(encryptedFile);
        }
    }
}
