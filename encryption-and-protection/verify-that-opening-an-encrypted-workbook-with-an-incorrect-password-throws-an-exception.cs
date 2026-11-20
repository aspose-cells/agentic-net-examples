// Title: How to verify that loading an encrypted Excel file with a wrong password raises a CellsException in Aspose.Cells for .NET
// Common Searches: Aspose.Cells catch CellsException when opening password‑protected XLSX with wrong password | C# test invalid password handling for encrypted workbook using Aspose.Cells LoadOptions | verify exception is thrown for incorrect password on protected Excel file in Aspose.Cells | unit test Aspose.Cells workbook load fails with wrong password | how to handle wrong password error when loading encrypted Excel in .NET
// Tags: load encrypted workbook with incorrect password Aspose.Cells | CellsException for invalid password | LoadOptions password validation Aspose.Cells | encrypted Excel file handling .NET | exception handling Aspose.Cells workbook load

using System;
using Aspose.Cells;

// The example creates a workbook, protects it with a password, saves it, and then attempts to open the file using LoadOptions with an incorrect password. It demonstrates that Aspose.Cells throws a CellsException for the invalid password and shows how to catch and differentiate expected versus unexpected exceptions.
class Program
{
    static void Main()
    {
        // Create a simple workbook and protect it with a password
        var workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].PutValue("Sample Data");
        workbook.Settings.Password = "correctPassword";

        string filePath = "encrypted.xlsx";
        workbook.Save(filePath); // Save the encrypted workbook

        // Prepare load options with an incorrect password
        var loadOptions = new LoadOptions(LoadFormat.Xlsx)
        {
            Password = "wrongPassword"
        };

        try
        {
            // Attempt to open the encrypted workbook using the wrong password
            var loadedWorkbook = new Workbook(filePath, loadOptions);
            // If no exception is thrown, the test has failed
            Console.WriteLine("Test Failed: No exception was thrown for an incorrect password.");
        }
        catch (CellsException ex)
        {
            // Expected path: Aspose.Cells throws a CellsException for invalid passwords
            Console.WriteLine($"Test Passed: Caught expected exception - {ex.Message}");
        }
        catch (Exception ex)
        {
            // Any other exception type is also considered a failure of the expectation
            Console.WriteLine($"Test Failed: Caught unexpected exception type {ex.GetType().Name} - {ex.Message}");
        }
    }
}
