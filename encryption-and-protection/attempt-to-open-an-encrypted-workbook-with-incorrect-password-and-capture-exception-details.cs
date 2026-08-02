// Title: Handle IncorrectPassword exception when opening an encrypted Excel file with Aspose.Cells in C#
// Description: Shows how to load an encrypted workbook using LoadOptions with an invalid password, catch the Aspose.Cells CellsException (ExceptionType.IncorrectPassword), display its message and code, and handle any other errors.
// Keywords: Aspose.Cells | C# | encrypted workbook | incorrect password | CellsException | ExceptionType.IncorrectPassword | LoadOptions password | Excel protection | exception handling | Workbook load error
// Common Searches: Aspose.Cells catch incorrect password exception | How to open password protected Excel with wrong password in C# | Exception thrown for wrong workbook password Aspose.Cells | LoadOptions password error handling | C# Aspose.Cells workbook load exception
// Developer Intent: Attempt to open a password‑protected workbook with a wrong password and retrieve the specific IncorrectPassword exception details.
// Use Cases: Validate user‑entered passwords and provide precise error feedback. | Log exception code and message for troubleshooting failed workbook loads. | Show a user‑friendly message when a protected workbook cannot be opened. | Implement fallback logic after detecting an incorrect password.
// AI Prompts: Write C# code using Aspose.Cells that opens an encrypted Excel workbook and distinguishes IncorrectPassword from other errors. | Explain how to extract the exception code and message from a CellsException when the password is wrong. | Show how to check if a workbook is encrypted before supplying a password with Aspose.Cells.

using System;
using Aspose.Cells;

// Shows how to load an encrypted workbook using LoadOptions with an invalid password, catch the Aspose.Cells CellsException (ExceptionType.IncorrectPassword), display its message and code, and handle any other errors.
class OpenEncryptedWorkbookDemo
{
    static void Main()
    {
        // Path to the encrypted workbook file
        string filePath = "encrypted.xlsx";

        // LoadOptions with an incorrect password
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.Password = "wrongPassword";

        try
        {
            // Attempt to open the workbook using the wrong password
            Workbook workbook = new Workbook(filePath, loadOptions);
            Console.WriteLine("Workbook opened successfully (unexpected).");
        }
        catch (CellsException ex) when (ex.Code == ExceptionType.IncorrectPassword)
        {
            // Capture specific Aspose.Cells exception for incorrect password
            Console.WriteLine("Incorrect password exception caught.");
            Console.WriteLine("Message: " + ex.Message);
            Console.WriteLine("Exception Code: " + ex.Code);
        }
        catch (Exception ex)
        {
            // Capture any other unexpected exceptions
            Console.WriteLine("General exception caught.");
            Console.WriteLine("Message: " + ex.Message);
        }
    }
}
