// Title: C# – Validate Excel Worksheet Protection Password with Aspose.Cells
// Description: Load an Excel workbook using Aspose.Cells, access the first worksheet, and use Worksheet.Protection.VerifyPassword to check a supplied password. Returns a Boolean indicating whether the password matches the sheet's protection password.
// Keywords: Aspose.Cells verify worksheet password C# | Excel sheet protection password check .NET | Worksheet.Protection.VerifyPassword example | validate Excel worksheet password programmatically | C# password validation for protected worksheet
// Common Searches: how to check worksheet password with Aspose.Cells | C# code to verify Excel sheet protection password | Aspose.Cells verify password for first worksheet | determine if Excel worksheet is password protected using .NET
// Developer Intent: Find out if a given string matches the protection password of the first worksheet in an Excel file.
// Use Cases: Gate editing access to a protected sheet until the correct password is supplied. | Validate uploaded workbooks in a web service before processing protected worksheets. | Automate batch jobs that skip or flag files whose first worksheet password does not match the expected value.
// AI Prompts: Write C# code that loads an Excel file with Aspose.Cells and returns true only when the supplied password matches the first worksheet's protection password. | Show how to handle exceptions when opening a workbook and verifying a worksheet password using Aspose.Cells. | Create a sample that iterates through all worksheets in a workbook and reports which ones are password‑protected and whether a given password is correct.

using System;
using Aspose.Cells;

// Load an Excel workbook using Aspose.Cells, access the first worksheet, and use Worksheet.Protection.VerifyPassword to check a supplied password. Returns a Boolean indicating whether the password matches the sheet's protection password.
public class WorksheetPasswordValidator
{
    // Validates the supplied password against the protection password of the first worksheet.
    // Returns true if the password matches, false otherwise.
    public static bool ValidateWorksheetPassword(string filePath, string password)
    {
        // Load the workbook from the specified file.
        Workbook workbook = new Workbook(filePath);
        // Access the first worksheet.
        Worksheet worksheet = workbook.Worksheets[0];

        // If the worksheet is not protected with a password, validation fails.
        if (!worksheet.Protection.IsProtectedWithPassword)
        {
            return false;
        }

        // Verify the supplied password using the Protection.VerifyPassword method.
        bool isValid = worksheet.Protection.VerifyPassword(password);
        return isValid;
    }

    // Example usage.
    public static void Main()
    {
        string filePath = "ProtectedSheet.xlsx";
        string correctPassword = "mySecret";
        string wrongPassword = "incorrect";

        bool isCorrect = ValidateWorksheetPassword(filePath, correctPassword);
        Console.WriteLine($"Correct password validation result: {isCorrect}");

        bool isWrong = ValidateWorksheetPassword(filePath, wrongPassword);
        Console.WriteLine($"Wrong password validation result: {isWrong}");
    }
}
