// Title: Test that an encrypted Excel workbook cannot be opened with an incorrect password using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that creates a password‑protected .xlsx file, then attempts to open it using LoadOptions with a wrong password and verifies that a CellsException is thrown. | Show how to implement a C# unit test that asserts loading an encrypted workbook with an invalid password does not expose any cell values.
// Common Searches: asp.net unit test for loading password protected Excel file with wrong password Aspose.Cells | how to catch CellsException when opening encrypted .xlsx with incorrect password in C# | verify that encrypted workbook prevents data access on invalid password using Aspose.Cells
// Tags: Aspose.Cells load encrypted workbook with wrong password | C# password protection Excel file Aspose.Cells | CellsException handling for invalid workbook password | unit testing workbook encryption Aspose.Cells | prevent data exposure encrypted XLSX Aspose.Cells

using System;
using Aspose.Cells;

// The example creates an .xlsx workbook, applies a password, saves it, then tries to load it with an incorrect password via LoadOptions. A CellsException is caught to confirm the workbook cannot be opened and no data is exposed, followed by cleanup of the temporary file.
class Program
{
    static void Main()
    {
        // Path for the temporary encrypted workbook
        string filePath = "encrypted.xlsx";

        // ------------------- Create -------------------
        // Create a new workbook and add some confidential data
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];
        ws.Cells["A1"].PutValue("Secret Data");

        // Protect the workbook with a password
        wb.Settings.Password = "correctPassword";

        // Save the encrypted workbook
        wb.Save(filePath, SaveFormat.Xlsx);

        // ------------------- Load (incorrect password) -------------------
        // Prepare load options with a wrong password
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
        loadOptions.Password = "wrongPassword";

        try
        {
            // Attempt to open the encrypted workbook using the wrong password
            Workbook wbWrong = new Workbook(filePath, loadOptions);

            // If no exception is thrown, the workbook was opened incorrectly
            Console.WriteLine("Test Failed: Workbook opened with an incorrect password.");
        }
        catch (CellsException)
        {
            // Expected outcome: an exception is thrown for a wrong password
            Console.WriteLine("Test Passed: Unable to open workbook with an incorrect password.");
            // No data is exposed because the exception prevents access to the workbook contents
        }
        finally
        {
            // Clean up the temporary file
            if (System.IO.File.Exists(filePath))
                System.IO.File.Delete(filePath);
        }
    }
}
