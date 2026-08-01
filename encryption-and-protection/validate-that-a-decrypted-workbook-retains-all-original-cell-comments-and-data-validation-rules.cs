// Title: Verify comments and list data validation persist after decrypting a password‑protected workbook with Aspose.Cells for .NET
// Description: Shows how to create an Excel file, add a comment to B2 and a list‑type validation to C3, protect it using Settings.Password, save it, reload with LoadOptions.Password, and programmatically confirm that both the comment and the validation remain unchanged.
// Keywords: Aspose.Cells | C# decrypt workbook | preserve cell comments | list data validation | password protected Excel | LoadOptions.Password | Settings.Password | .NET | Excel encryption verification | cell comment validation | data validation integrity
// Common Searches: Aspose.Cells keep comments after opening encrypted file | check data validation after decrypting Excel with Aspose.Cells | load password protected workbook .NET Aspose | verify cell comment author after decryption | ensure list validation persists in encrypted workbook
// Developer Intent: Confirm that a workbook opened with the correct password still contains the original comment and list validation.
// Use Cases: Automated test to validate integrity of protected Excel files in CI pipelines. | Utility method that returns true when the comment at B2 matches expected text and author after decryption. | Audit script that compares original and decrypted workbooks for comment and validation consistency. | Migration tool that verifies data validation rules survive encryption. | Batch processor that reports missing comments or validations in password‑protected files.
// AI Prompts: Create a reusable C# function that opens a password‑protected Excel file with Aspose.Cells and checks if the comment at B2 matches given text and author. | Write an xUnit test that builds a workbook, adds a comment and list validation, encrypts it, decrypts it, and asserts both features are unchanged. | Generate code to log detailed mismatches when comment note, author, or validation formula differ after decryption. | Provide a PowerShell script that uses Aspose.Cells to process multiple encrypted files and report any missing comments or validations. | Explain how to compare two workbooks (original vs decrypted) for comment and validation equality using Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsValidationDemo
{
    // Shows how to create an Excel file, add a comment to B2 and a list‑type validation to C3, protect it using Settings.Password, save it, reload with LoadOptions.Password, and programmatically confirm that both the comment and the validation remain unchanged.
    class Program
    {
        static void Main()
        {
            // ---------- Create original workbook ----------
            Workbook originalWb = new Workbook();
            Worksheet sheet = originalWb.Worksheets[0];

            // Add a comment to cell B2
            int commentIdx = sheet.Comments.Add("B2");
            sheet.Comments[commentIdx].Note = "Original comment text";
            sheet.Comments[commentIdx].Author = "OriginalAuthor";

            // Add a list data validation to cell C3
            Validation validation = sheet.Validations[sheet.Validations.Add()];
            validation.Type = ValidationType.List;
            validation.Formula1 = "\"OptionA,OptionB,OptionC\""; // list items
            CellArea area = new CellArea
            {
                StartRow = 2,   // C3 -> row index 2
                StartColumn = 2,
                EndRow = 2,
                EndColumn = 2
            };
            validation.AddArea(area);

            // Protect workbook with a password
            originalWb.Settings.Password = "SecretPwd";

            // Save the encrypted workbook
            string encryptedPath = "EncryptedWorkbook.xlsx";
            originalWb.Save(encryptedPath);

            // ---------- Load decrypted workbook ----------
            LoadOptions loadOpts = new LoadOptions
            {
                Password = "SecretPwd"
            };
            Workbook decryptedWb = new Workbook(encryptedPath, loadOpts);
            Worksheet decryptedSheet = decryptedWb.Worksheets[0];

            // ---------- Validate comments ----------
            bool commentExists = decryptedSheet.Comments.Count > 0;
            bool commentMatches = false;
            if (commentExists)
            {
                // Assuming the comment is still at B2 (index 0)
                var loadedComment = decryptedSheet.Comments[0];
                commentMatches = loadedComment.Note == "Original comment text"
                                 && loadedComment.Author == "OriginalAuthor";
            }

            // ---------- Validate data validation ----------
            bool validationExists = decryptedSheet.Validations.Count > 0;
            bool validationMatches = false;
            if (validationExists)
            {
                // Retrieve validation for cell C3
                Validation loadedValidation = decryptedSheet.Validations.GetValidationInCell(2, 2);
                if (loadedValidation != null)
                {
                    validationMatches = loadedValidation.Type == ValidationType.List
                                        && loadedValidation.Formula1 == "\"OptionA,OptionB,OptionC\"";
                }
            }

            // ---------- Output results ----------
            Console.WriteLine("Comment exists: " + commentExists);
            Console.WriteLine("Comment matches original: " + commentMatches);
            Console.WriteLine("Data validation exists: " + validationExists);
            Console.WriteLine("Data validation matches original: " + validationMatches);
        }
    }
}
