// Title: Aspose.Cells .NET – Open Encrypted Excel Workbook Without Password and Capture IncorrectPassword Error Code
// Description: Demonstrates how to load a password‑protected XLSX file with Aspose.Cells' Workbook constructor, catch the CellsException when ExceptionType.IncorrectPassword is thrown, output the numeric error code, and handle any other unexpected errors.
// Keywords: Aspose.Cells open encrypted workbook | C# IncorrectPassword exception | CellsException error code | read password‑protected Excel file .NET | Aspose.Cells workbook without password | Excel encryption handling Aspose | ExceptionType.IncorrectPassword | capture numeric error code Aspose.Cells
// Common Searches: Aspose.Cells catch IncorrectPassword when opening protected Excel | Get error code for missing password in Aspose.Cells Workbook | How to detect encrypted Excel file with Aspose.Cells C# | Exception thrown for wrong password in Aspose.Cells | Retrieve numeric code for IncorrectPassword exception
// Developer Intent: The developer wants to try opening an encrypted Excel workbook without supplying a password, capture the specific IncorrectPassword exception, and read its numeric error code.
// Use Cases: Check if a file is password‑protected before prompting the user by attempting to open it and detecting ExceptionType.IncorrectPassword. | Log detailed troubleshooting data, including the exact error code, when an encrypted workbook is accessed without credentials. | Trigger a fallback workflow that requests the correct password or skips processing after catching the IncorrectPassword code.
// AI Prompts: Write C# code using Aspose.Cells to open an Excel file, detect a missing or wrong password, and print the numeric IncorrectPassword error code. | Show how to handle CellsException with ExceptionType.IncorrectPassword in Aspose.Cells, log the exception code, and re‑throw a custom exception. | Explain how to programmatically determine whether an Excel workbook is encrypted with Aspose.Cells without providing a password and retrieve the associated error code.

using System;
using Aspose.Cells;

// Demonstrates how to load a password‑protected XLSX file with Aspose.Cells' Workbook constructor, catch the CellsException when ExceptionType.IncorrectPassword is thrown, output the numeric error code, and handle any other unexpected errors.
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
        catch (CellsException ex) when (ex.Code == ExceptionType.IncorrectPassword)
        {
            // Capture the specific error code for an incorrect or missing password
            Console.WriteLine($"Incorrect password error captured. Exception code: {(int)ex.Code}");
        }
        catch (Exception ex)
        {
            // Handle any other unexpected exceptions
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
