// Title: How to verify that opening a password‑protected Excel workbook with an incorrect password throws a CellsException in Aspose.Cells for .NET
// AI Prompts: Write C# code that creates an Excel workbook, encrypts it with a password using Aspose.Cells, saves it, then attempts to open it with a different password and catches the resulting CellsException. | Generate a C# unit test using Aspose.Cells that asserts a CellsException is thrown when LoadOptions.Password does not match the workbook’s encryption password. | Provide a step‑by‑step example showing how to handle an invalid password error when loading an encrypted .xlsx file with Aspose.Cells in a .NET application.
// Common Searches: Aspose.Cells throws exception when opening encrypted .xlsx with wrong password | C# load password protected Excel file using Aspose.Cells and catch CellsException | How to test invalid workbook password handling in Aspose.Cells .NET | LoadOptions.Password incorrect value exception Aspose.Cells example | Validate encryption password mismatch error with Aspose.Cells for .NET
// Tags: Aspose.Cells load encrypted workbook exception | LoadOptions incorrect password handling | CellsException invalid password | C# workbook encryption validation | Aspose.Cells password protection error

using System;
using Aspose.Cells;

// The example creates a workbook, sets a password, saves it as an encrypted .xlsx file, then attempts to open the file with a mismatched password using LoadOptions, catching the CellsException to confirm that Aspose.Cells correctly reports an invalid password error.
class Program
{
    static void Main()
    {
        // Create a new workbook and add sample data
        Workbook wb = new Workbook();
        wb.Worksheets[0].Cells["A1"].PutValue("Sample");

        // Encrypt the workbook with a known password
        wb.Settings.Password = "correctPassword";
        string filePath = "encrypted.xlsx";
        wb.Save(filePath);

        // Attempt to open the encrypted workbook using an incorrect password
        try
        {
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
            loadOptions.Password = "wrongPassword";

            // This line should throw an exception because the password is incorrect
            Workbook wbWrong = new Workbook(filePath, loadOptions);

            // If no exception is thrown, the test has failed
            Console.WriteLine("Test Failed: No exception was thrown.");
        }
        catch (Aspose.Cells.CellsException ex)
        {
            // Expected path: an exception is thrown for an invalid password
            Console.WriteLine("Test Passed: Exception thrown as expected.");
            Console.WriteLine("Exception Message: " + ex.Message);
        }
    }
}
