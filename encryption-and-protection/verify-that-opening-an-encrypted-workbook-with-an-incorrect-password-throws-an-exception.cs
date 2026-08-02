// Title: Aspose.Cells for .NET – Verify exception when opening an encrypted workbook with an incorrect password
// Description: Creates a workbook, encrypts it with a password, saves it, then attempts to load the file using a wrong password via LoadOptions. The code demonstrates that Aspose.Cells throws an exception for an invalid password and confirms successful opening with the correct password.
// Keywords: Aspose.Cells | C# | .NET | encrypted workbook | wrong password | LoadOptions | exception handling | password protection | Excel encryption
// Common Searches: Aspose.Cells open encrypted Excel with wrong password | C# catch exception for invalid workbook password Aspose.Cells | How to test password protection in Aspose.Cells .NET | LoadOptions password exception Aspose.Cells | Verify encrypted workbook access Aspose.Cells
// Developer Intent: Confirm that loading a password‑protected workbook with an invalid password raises an exception in Aspose.Cells for .NET.
// Use Cases: Validate that workbook encryption blocks unauthorized access by catching the thrown exception. | Automated test to ensure password protection behaves as expected before deployment. | Log failed password attempts by capturing exception details for security auditing.
// AI Prompts: Generate an xUnit test that asserts Aspose.Cells throws a specific exception when opening an encrypted workbook with an incorrect password. | Show how to catch the Aspose.Cells.PasswordIncorrectException and retrieve its error code. | Explain how to programmatically detect whether an Excel file is password‑protected before attempting to open it with Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsPasswordVerification
{
    // Creates a workbook, encrypts it with a password, saves it, then attempts to load the file using a wrong password via LoadOptions. The code demonstrates that Aspose.Cells throws an exception for an invalid password and confirms successful opening with the correct password.
    class Program
    {
        static void Main()
        {
            // Step 1: Create a new workbook and add some data
            Workbook wb = new Workbook();
            Worksheet sheet = wb.Worksheets[0];
            sheet.Cells["A1"].PutValue("Encrypted content");

            // Step 2: Set a password to encrypt the workbook
            wb.Settings.Password = "correctPassword";

            // Step 3: Save the encrypted workbook
            string encryptedFile = "encrypted.xlsx";
            wb.Save(encryptedFile);

            // Step 4: Attempt to open the encrypted workbook with an incorrect password
            LoadOptions wrongOptions = new LoadOptions();
            wrongOptions.Password = "wrongPassword";

            try
            {
                // This should throw an exception because the password is incorrect
                Workbook wbWrong = new Workbook(encryptedFile, wrongOptions);
                Console.WriteLine("Unexpectedly opened workbook with wrong password.");
            }
            catch (Exception ex)
            {
                // Expected path: an exception is thrown
                Console.WriteLine("Failed to open workbook with incorrect password as expected.");
                Console.WriteLine("Exception message: " + ex.Message);
            }

            // Optional: Verify that opening with the correct password succeeds
            LoadOptions correctOptions = new LoadOptions();
            correctOptions.Password = "correctPassword";

            try
            {
                Workbook wbCorrect = new Workbook(encryptedFile, correctOptions);
                Console.WriteLine("Workbook opened successfully with correct password.");
                Console.WriteLine("Cell A1 value: " + wbCorrect.Worksheets[0].Cells["A1"].Value);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed to open workbook with correct password.");
                Console.WriteLine("Exception message: " + ex.Message);
            }
        }
    }
}
