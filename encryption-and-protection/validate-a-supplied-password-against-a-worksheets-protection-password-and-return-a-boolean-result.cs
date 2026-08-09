// Title: Validate Worksheet Protection Password with Aspose.Cells for .NET (C#)
// Description: Loads an Excel workbook, accesses a worksheet, and uses Aspose.Cells' Worksheet.Protection.VerifyPassword method to determine if a supplied password matches the sheet's protection password, returning a boolean without modifying the file.
// Keywords: Aspose.Cells | C# | .NET | worksheet protection | verify password | Protection.VerifyPassword | Excel sheet password validation | validate worksheet password | check worksheet protection | Aspose.Cells example | Excel security
// Common Searches: Aspose.Cells verify worksheet password C# | How to check Excel sheet protection password using Aspose.Cells | C# code to validate worksheet protection password | Worksheet.Protection.VerifyPassword sample | Validate Excel sheet password without opening UI
// Developer Intent: Determine if a supplied string matches the protection password of a specific worksheet in an Excel workbook using Aspose.Cells.
// Use Cases: Prompt a user for a password and enable editing only when the worksheet unlocks. | Scan multiple workbooks to flag sheets that are protected with a known password before automated processing. | Expose a REST endpoint that receives a file path and password, returning true/false to indicate worksheet access.
// AI Prompts: Generate C# code with Aspose.Cells that validates a worksheet's protection password and includes error handling for missing files or unprotected sheets. | Show how to verify the password of the second worksheet (index 1) instead of the first one using Aspose.Cells. | Explain how to retrieve the hashed protection password from a worksheet with Aspose.Cells and compare it manually to a user‑provided password.

using System;
using Aspose.Cells;

namespace AsposeCellsPasswordValidation
{
    // Loads an Excel workbook, accesses a worksheet, and uses Aspose.Cells' Worksheet.Protection.VerifyPassword method to determine if a supplied password matches the sheet's protection password, returning a boolean without modifying the file.
    public class WorksheetPasswordValidator
    {
        /// <param name="filePath">Path to the Excel file.</param>
        /// <param name="password">Password to validate.</param>
        /// <returns>True if the password is correct; otherwise, false.</returns>
        public static bool ValidateWorksheetPassword(string filePath, string password)
        {
            // Load the workbook (creation/load rule)
            Workbook workbook = new Workbook(filePath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Verify the supplied password against the worksheet's protection password
            // Uses Protection.VerifyPassword method as defined in the documentation
            bool isValid = worksheet.Protection.VerifyPassword(password);

            // No need to modify or save the workbook for validation purposes
            return isValid;
        }

        // Example usage
        public static void Main()
        {
            // Path to an existing workbook that has worksheet protection enabled
            string filePath = "ProtectedWorksheet.xlsx";

            // Password to test
            string testPassword = "password123";

            bool result = ValidateWorksheetPassword(filePath, testPassword);
            Console.WriteLine($"Password validation result: {result}");
        }
    }
}
