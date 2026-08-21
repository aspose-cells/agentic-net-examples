// Title: Handle IncorrectPassword Exception when Opening an Encrypted Excel Workbook with Aspose.Cells for .NET (C#)
// Description: Shows how to create a password‑protected workbook, save it, then attempt to load it with a wrong password using LoadOptions. The code catches CellsException with ExceptionType.IncorrectPassword, prints the message and code, and handles any other exceptions separately.
// Keywords: Aspose.Cells | C# | encrypted workbook | incorrect password | CellsException | ExceptionType.IncorrectPassword | LoadOptions password | Excel encryption .NET | catch password exception | open protected workbook
// Common Searches: Aspose.Cells catch incorrect password | How to handle CellsException IncorrectPassword | Load encrypted Excel file with wrong password Aspose.Cells | Get exception details for wrong password Aspose.Cells | C# open password protected workbook Aspose.Cells
// Developer Intent: Attempt to open a password‑protected Excel file with an invalid password and capture the specific IncorrectPassword exception details.
// Use Cases: Validate a user‑entered password, log audit information, and prevent unauthorized access when the password is wrong. | Display a friendly error message in a UI when Aspose.Cells reports an incorrect password. | Write unit tests that confirm Aspose.Cells throws CellsException with code IncorrectPassword for invalid credentials.
// AI Prompts: Generate C# code that opens an encrypted workbook using Aspose.Cells, accepts a password argument, and returns a structured error object when the password is invalid. | Create a NUnit test that verifies Aspose.Cells throws CellsException with ExceptionType.IncorrectPassword when loading a workbook with an incorrect password. | Refactor the example to log the exception message and code to a file and rethrow a custom exception for higher‑level handling.

using System;
using Aspose.Cells;

// Shows how to create a password‑protected workbook, save it, then attempt to load it with a wrong password using LoadOptions. The code catches CellsException with ExceptionType.IncorrectPassword, prints the message and code, and handles any other exceptions separately.
class Program
{
    static void Main()
    {
        // Create a new workbook and add some data
        Workbook wb = new Workbook();
        wb.Worksheets[0].Cells["A1"].PutValue("Secret Data");

        // Protect the workbook with a password
        wb.Settings.Password = "correctPassword";

        // Save the encrypted workbook
        string filePath = "encryptedWorkbook.xlsx";
        wb.Save(filePath);

        // Prepare load options with an incorrect password
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.Password = "wrongPassword";

        try
        {
            // Attempt to open the encrypted workbook using the wrong password
            Workbook wbWrong = new Workbook(filePath, loadOptions);
            Console.WriteLine("Workbook opened successfully (unexpected).");
        }
        catch (CellsException ex) when (ex.Code == ExceptionType.IncorrectPassword)
        {
            // Capture and display details of the incorrect password exception
            Console.WriteLine("Incorrect password exception caught:");
            Console.WriteLine($"Message: {ex.Message}");
            Console.WriteLine($"Exception Code: {ex.Code}");
        }
        catch (Exception ex)
        {
            // Capture any other unexpected exceptions
            Console.WriteLine("An unexpected exception occurred:");
            Console.WriteLine(ex.Message);
        }
    }
}
