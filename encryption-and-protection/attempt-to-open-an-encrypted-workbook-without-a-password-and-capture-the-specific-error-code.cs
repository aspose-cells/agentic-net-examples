// Title: Capture Aspose.Cells IncorrectPassword code when opening an encrypted workbook without a password (C#)
// Description: A C# example that tries to load a password‑protected Excel file with Aspose.Cells, catches the CellsException, and prints the specific error code (ExceptionType.IncorrectPassword = 8) that indicates a missing or wrong password.
// Keywords: Aspose.Cells | encrypted workbook | password protection | IncorrectPassword | CellsException | exception code | C# | open Excel file | missing password | error handling
// Common Searches: Aspose.Cells error code for missing password | How to catch IncorrectPassword exception in Aspose.Cells | Open password‑protected Excel file without password Aspose.Cells | Retrieve exception code when opening encrypted workbook | What does exception code 8 mean in Aspose.Cells
// Developer Intent: Identify the exact Aspose.Cells exception code returned when an encrypted Excel workbook is opened without providing a password.
// Use Cases: Detect a missing or wrong password and show a user‑friendly prompt. | Log the precise CellsException code for security auditing. | Trigger conditional logic to request a password only when the IncorrectPassword code is received. | Integrate password validation into automated file‑processing pipelines.
// AI Prompts: Generate C# code that opens an Excel file with Aspose.Cells, catches CellsException, and outputs the IncorrectPassword code when no password is supplied. | Explain how to differentiate IncorrectPassword from other Aspose.Cells error codes in a try‑catch block. | Write a method that returns the Aspose.Cells exception code when opening a protected workbook fails. | Create a unit test that verifies the IncorrectPassword code is returned for an encrypted file opened without a password.

using System;
using Aspose.Cells;

// A C# example that tries to load a password‑protected Excel file with Aspose.Cells, catches the CellsException, and prints the specific error code (ExceptionType.IncorrectPassword = 8) that indicates a missing or wrong password.
class Program
{
    static void Main()
    {
        // Path to the encrypted workbook
        string filePath = "encrypted.xlsx";

        try
        {
            // Attempt to open the workbook without providing a password
            Workbook workbook = new Workbook(filePath);
            Console.WriteLine("Workbook opened successfully (this is unexpected for an encrypted file).");
        }
        catch (CellsException ex)
        {
            // Capture and display the specific Aspose.Cells error code
            // For an incorrect or missing password the code is ExceptionType.IncorrectPassword (value 8)
            Console.WriteLine($"Failed to open workbook. Exception code: {ex.Code}");
        }
        catch (Exception ex)
        {
            // Handle any other unexpected exceptions
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
