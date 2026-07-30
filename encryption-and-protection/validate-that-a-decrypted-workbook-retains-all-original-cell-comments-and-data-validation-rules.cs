// Title: C# – Verify that decrypting a password‑protected Excel file keeps cell comments and data‑validation rules (Aspose.Cells)
// Description: This .NET example creates a workbook, adds a comment to A1 and a whole‑number validation to B2, encrypts the file with a password, saves it, then reloads it using LoadOptions.Password. After decryption it confirms that the comment count, note and author are unchanged and that the validation’s type, operator and formulas match the original settings.
// Keywords: Aspose.Cells | C# | .NET | Excel encryption | password‑protected workbook | cell comments | data validation | LoadOptions.Password | workbook decryption | preserve Excel metadata
// Common Searches: How to keep cell comments after encrypting an Excel file with Aspose.Cells | Verify data validation survives workbook password protection in C# | Load a password‑protected XLSX and check comments using Aspose.Cells | Aspose.Cells example for preserving validation rules after encryption | C# code to decrypt an Excel file and confirm original metadata
// Developer Intent: Confirm that a workbook opened after decryption still contains the exact comments and validation rules defined before encryption.
// Use Cases: Open an encrypted workbook with LoadOptions.Password and assert that ws.Comments[0].Note and Author equal the pre‑encryption values. | Iterate ws.Validations after decryption to verify Type, Operator, Formula1 and Formula2 match the original whole‑number rule. | Add integrity checks to CI pipelines that automatically validate comment and validation retention for protected Excel files. | Create a utility method that returns true only when all original comments and data‑validation objects are present after decryption.
// AI Prompts: Write a C# method that accepts an encrypted Excel path and password, opens it with Aspose.Cells, and returns a boolean indicating whether every original comment and data‑validation rule is intact. | Generate an NUnit test that encrypts a workbook containing a comment and a whole‑number validation, then decrypts it and asserts that the comment text, author, and validation properties are unchanged. | Provide a step‑by‑step tutorial for verifying that data‑validation rules are preserved after saving an encrypted workbook using Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsValidationDemo
{
    // This .NET example creates a workbook, adds a comment to A1 and a whole‑number validation to B2, encrypts the file with a password, saves it, then reloads it using LoadOptions.Password. After decryption it confirms that the comment count, note and author are unchanged and that the validation’s type, operator and formulas match the original settings.
    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a new workbook and add a comment and a data validation
            // -----------------------------------------------------------------
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Add a comment to cell A1
            int commentIndex = ws.Comments.Add("A1");
            ws.Comments[commentIndex].Note = "Original comment text";
            ws.Comments[commentIndex].Author = "OriginalAuthor";

            // Add a whole-number data validation to cell B2 (range 10-100)
            Validation validation = ws.Validations[ws.Validations.Add()];
            validation.Type = ValidationType.WholeNumber;
            validation.Operator = OperatorType.Between;
            validation.Formula1 = "10";
            validation.Formula2 = "100";

            CellArea area = new CellArea
            {
                StartRow = 1,   // B2 row index (0‑based)
                StartColumn = 1,
                EndRow = 1,
                EndColumn = 1
            };
            validation.AddArea(area);

            // --------------------------------------------------------------
            // 2. Encrypt the workbook with a password and save it to disk
            // --------------------------------------------------------------
            string password = "SecretPwd123";
            wb.Settings.Password = password;               // encrypt
            string encryptedPath = "EncryptedWorkbook.xlsx";
            wb.Save(encryptedPath);                         // save

            // --------------------------------------------------------------
            // 3. Load the encrypted workbook using the password (decryption)
            // --------------------------------------------------------------
            LoadOptions loadOptions = new LoadOptions
            {
                Password = password
            };
            Workbook decryptedWb = new Workbook(encryptedPath, loadOptions);
            Worksheet decryptedWs = decryptedWb.Worksheets[0];

            // --------------------------------------------------------------
            // 4. Validate that the comment is retained
            // --------------------------------------------------------------
            bool commentExists = decryptedWs.Comments.Count > 0;
            string commentNote = commentExists ? decryptedWs.Comments[0].Note : string.Empty;
            string commentAuthor = commentExists ? decryptedWs.Comments[0].Author : string.Empty;

            Console.WriteLine("Comment exists: " + commentExists);
            Console.WriteLine("Comment text matches: " + (commentNote == "Original comment text"));
            Console.WriteLine("Comment author matches: " + (commentAuthor == "OriginalAuthor"));

            // --------------------------------------------------------------
            // 5. Validate that the data validation rule is retained
            // --------------------------------------------------------------
            bool validationExists = decryptedWs.Validations.Count > 0;
            Validation decryptedValidation = validationExists ? decryptedWs.Validations[0] : null;

            bool validationMatches = false;
            if (validationExists && decryptedValidation != null)
            {
                // Check type, operator and formulas
                validationMatches =
                    decryptedValidation.Type == ValidationType.WholeNumber &&
                    decryptedValidation.Operator == OperatorType.Between &&
                    decryptedValidation.Formula1 == "10" &&
                    decryptedValidation.Formula2 == "100";
            }

            Console.WriteLine("Data validation exists: " + validationExists);
            Console.WriteLine("Data validation matches original: " + validationMatches);
        }
    }
}
