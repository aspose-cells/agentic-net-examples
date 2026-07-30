// Title: Catch IncorrectPassword error when opening an encrypted Excel file with Aspose.Cells (C#)
// Description: Demonstrates loading an encrypted workbook without a password, catching the CellsException, reading its error code, and confirming it matches ExceptionType.IncorrectPassword to identify password protection.
// Keywords: Aspose.Cells | encrypted workbook | IncorrectPassword | CellsException | C# | password protection | error code | Workbook constructor | exception handling
// Common Searches: Aspose.Cells error code for encrypted workbook | How to detect password‑protected Excel file using Aspose.Cells | IncorrectPassword exception Aspose.Cells .NET | Open encrypted Excel without password Aspose.Cells | Retrieve CellsException.Code in C#
// Developer Intent: Identify the exact Aspose.Cells exception type and code returned when a password‑protected Excel file is opened without supplying a password.
// Use Cases: Programmatically verify whether a file is encrypted before prompting for a password. | Log the specific error code for audit trails when unauthorized access to a protected workbook occurs. | Implement conditional logic that only asks the user for credentials after detecting ExceptionType.IncorrectPassword. | Integrate encryption detection into a batch file‑validation workflow.
// AI Prompts: Generate C# code using Aspose.Cells that attempts to open a workbook, catches CellsException, and returns true if ExceptionType.IncorrectPassword is detected. | Create a robust error‑handling routine for Aspose.Cells that distinguishes IncorrectPassword, FileNotFound, and other exception types, logging each error code. | Show how to extract and display the numeric value of CellsException.Code when opening a password‑protected Excel file with Aspose.Cells.

using System;
using Aspose.Cells;

// Demonstrates loading an encrypted workbook without a password, catching the CellsException, reading its error code, and confirming it matches ExceptionType.IncorrectPassword to identify password protection.
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
            Console.WriteLine("Workbook opened successfully (unexpected).");
        }
        catch (CellsException ex)
        {
            // Capture and display the specific Aspose.Cells error code
            Console.WriteLine($"Caught CellsException. Error code: {ex.Code}");

            // Verify that the error corresponds to an incorrect password
            if (ex.Code == ExceptionType.IncorrectPassword)
            {
                Console.WriteLine("The workbook is encrypted and requires a password (IncorrectPassword).");
            }
        }
    }
}
