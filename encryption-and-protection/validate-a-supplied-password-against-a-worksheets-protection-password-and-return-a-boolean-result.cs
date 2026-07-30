// Title: C# – Verify an Excel worksheet’s protection password with Aspose.Cells
// Description: Load an existing workbook, access a worksheet, and use Aspose.Cells Worksheet.Protection.VerifyPassword to compare a supplied password with the worksheet’s protection password, returning a Boolean result.
// Keywords: Aspose.Cells worksheet password verification | C# Excel protection password check | Verify worksheet protection Aspose | Protection.VerifyPassword example | Aspose.Cells password validation .NET
// Common Searches: how to check worksheet password using Aspose.Cells C# | Aspose.Cells verify worksheet protection password example | C# code to validate Excel sheet password without opening Excel | Aspose.Cells Protection.VerifyPassword usage
// Developer Intent: Determine whether a given string matches the protection password of a worksheet in an Excel file using Aspose.Cells for .NET.
// Use Cases: Prompt users for a password before unlocking a protected sheet in a desktop or web application. | Skip or flag workbooks in a batch job when the worksheet password cannot be confirmed. | Expose a REST endpoint that validates uploaded Excel files by confirming worksheet protection passwords.
// AI Prompts: Write C# code that uses Aspose.Cells to verify a worksheet's protection password and gracefully handles missing files or incorrect passwords. | Show how to loop through all worksheets in a workbook and validate each protection password with Aspose.Cells. | Create a sample that logs the protection status of a worksheet after calling VerifyPassword, including exception handling.

using System;
using Aspose.Cells;

namespace AsposeCellsPasswordValidation
{
    // Load an existing workbook, access a worksheet, and use Aspose.Cells Worksheet.Protection.VerifyPassword to compare a supplied password with the worksheet’s protection password, returning a Boolean result.
    public class WorksheetPasswordValidator
    {
        /// <param name="workbookPath">Full path to the Excel file.</param>
        /// <param name="password">Password to validate.</param>
        /// <returns>Boolean indicating whether the password is correct.</returns>
        public static bool Validate(string workbookPath, string password)
        {
            // Load the workbook (existing file should already have worksheet protection set)
            Workbook workbook = new Workbook(workbookPath);

            // Access the first worksheet (index 0)
            Worksheet worksheet = workbook.Worksheets[0];

            // Use the Protection.VerifyPassword method to check the password
            bool isPasswordCorrect = worksheet.Protection.VerifyPassword(password);

            // Return the verification result
            return isPasswordCorrect;
        }

        // Example usage
        public static void Main()
        {
            // Path to the workbook that has its first worksheet protected with a password
            string filePath = "ProtectedWorksheet.xlsx";

            // Password to test
            string testPassword = "password123";

            // Perform validation
            bool result = Validate(filePath, testPassword);

            Console.WriteLine($"Password validation result: {result}");
        }
    }
}
