// Title: How to verify that loading an encrypted Excel workbook with a wrong password raises a CellsException using Aspose.Cells for .NET
// Common Searches: Aspose.Cells catch exception for invalid password when opening encrypted .xlsx in C# | C# test loading password‑protected workbook with wrong password using LoadOptions | How to handle CellsException for incorrect workbook password in Aspose.Cells | Validate encrypted workbook opening fails with wrong password in .NET | Load encrypted Excel file with incorrect password Aspose.Cells exception handling
// Tags: encrypted workbook loading exception Aspose.Cells | LoadOptions wrong password handling | CellsException invalid password detection | C# Aspose.Cells password protection error | opening password‑protected .xlsx with incorrect password

using System;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, encrypts it with a password, saves it to a memory stream, then attempts to open it using LoadOptions set to an incorrect password. It demonstrates that Aspose.Cells throws a CellsException for an invalid password and shows proper exception handling in C#.
class Program
{
    static void Main()
    {
        // Create a new workbook and set a password for encryption
        Workbook workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].PutValue("Sample Data");
        workbook.Settings.Password = "correctPassword";

        // Save the encrypted workbook to a memory stream
        using (MemoryStream stream = new MemoryStream())
        {
            workbook.Save(stream, SaveFormat.Xlsx);
            stream.Position = 0; // Reset stream position for reading

            // Prepare load options with an incorrect password
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = "wrongPassword"
            };

            try
            {
                // Attempt to load the workbook using the wrong password
                Workbook loadedWorkbook = new Workbook(stream, loadOptions);
                // If no exception is thrown, the test has failed
                Console.WriteLine("Test Failed: No exception was thrown for an incorrect password.");
            }
            catch (CellsException ex)
            {
                // Expected outcome: an exception indicating an invalid password
                Console.WriteLine($"Test Passed: Caught expected exception - {ex.Message}");
            }
            catch (Exception ex)
            {
                // Any other exception type is unexpected
                Console.WriteLine($"Test Failed: Unexpected exception type - {ex.GetType().Name}: {ex.Message}");
            }
        }
    }
}
