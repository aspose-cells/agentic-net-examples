// Title: C# – Verify that comments and data‑validation rules persist after decrypting an Excel workbook with Aspose.Cells
// Description: This example creates a workbook, adds a comment to A1 and a whole‑number validation to A2, encrypts the file with a password, saves it, then reloads it using LoadOptions.Password. After decryption it confirms that the comment’s note and author and the validation’s type, operator, and formulas are unchanged.
// Keywords: Aspose.Cells decrypt workbook | preserve Excel comments after encryption | data validation after password protection | C# load encrypted Excel file | verify workbook integrity Aspose | LoadOptions.Password example | Excel comment validation C#
// Common Searches: how to keep comments when opening a password‑protected Excel file with Aspose.Cells | check data validation after decrypting an encrypted workbook in .NET | compare original and decrypted workbook comments Aspose | verify validation rules survive Excel encryption C#
// Developer Intent: Ensure that a workbook opened with the correct password still contains the exact comment and data‑validation settings that were applied before encryption.
// Use Cases: Load a password‑protected workbook and assert that the first comment’s Note and Author match the expected values. | Retrieve the validation for cell A2 with GetValidationInCell and confirm that Type, Operator, Formula1, and Formula2 are identical to the original configuration. | Add integrity checks for comments and validations to an automated test suite for protected Excel files.
// AI Prompts: Write C# code that opens a password‑protected Excel file using Aspose.Cells and asserts that all comments and data‑validation rules are identical to those in the source workbook. | Show how to log detailed differences when a comment’s text, author, or a validation’s properties differ after loading an encrypted workbook. | Create a reusable method that takes the original and decrypted Workbook objects and returns true only if every comment and validation is preserved.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // This example creates a workbook, adds a comment to A1 and a whole‑number validation to A2, encrypts the file with a password, saves it, then reloads it using LoadOptions.Password. After decryption it confirms that the comment’s note and author and the validation’s type, operator, and formulas are unchanged.
    public class DecryptedWorkbookValidationDemo
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // -------------------- Create and configure workbook --------------------
            // Create a new workbook (lifecycle: create)
            using (Workbook originalWorkbook = new Workbook())
            {
                // Access the first worksheet
                Worksheet sheet = originalWorkbook.Worksheets[0];

                // ---------- Add a comment ----------
                // Add a comment to cell A1 and set its properties
                int commentIndex = sheet.Comments.Add("A1");
                sheet.Comments[commentIndex].Note = "Original comment text";
                sheet.Comments[commentIndex].Author = "Original Author";

                // ---------- Add a data validation ----------
                // Define the cell range for the validation (A2)
                CellArea area = new CellArea
                {
                    StartRow = 1,
                    StartColumn = 0,
                    EndRow = 1,
                    EndColumn = 0
                };

                // Create a validation that allows whole numbers between 10 and 100
                Validation validation = sheet.Validations[sheet.Validations.Add(area)];
                validation.Type = ValidationType.WholeNumber;
                validation.Operator = OperatorType.Between;
                validation.Formula1 = "10";
                validation.Formula2 = "100";

                // ---------- Encrypt the workbook ----------
                // Set a password to protect the workbook (encryption)
                originalWorkbook.Settings.Password = "SecretPwd";

                // Save the encrypted workbook (lifecycle: save)
                string encryptedPath = "EncryptedWorkbook.xlsx";
                originalWorkbook.Save(encryptedPath);
            }

            // -------------------- Load the encrypted workbook (decrypted) --------------------
            string encryptedFilePath = "EncryptedWorkbook.xlsx";

            if (!File.Exists(encryptedFilePath))
            {
                Console.WriteLine($"File not found: {encryptedFilePath}");
                return;
            }

            // Prepare load options with the password
            LoadOptions loadOptions = new LoadOptions
            {
                Password = "SecretPwd"
            };

            // Load the workbook using the password (lifecycle: load)
            using (Workbook loadedWorkbook = new Workbook(encryptedFilePath, loadOptions))
            {
                Worksheet loadedSheet = loadedWorkbook.Worksheets[0];

                // -------------------- Validate comments --------------------
                Comment loadedComment = null;
                bool commentExists = false;
                string commentNote = null;
                string commentAuthor = null;

                if (loadedSheet.Comments.Count > 0)
                {
                    // Assuming the comment we added is the first one
                    loadedComment = loadedSheet.Comments[0];
                    commentExists = loadedComment != null;
                    if (commentExists)
                    {
                        commentNote = loadedComment.Note;
                        commentAuthor = loadedComment.Author;
                    }
                }

                // -------------------- Validate data validation --------------------
                // Retrieve validation for cell A2 (row 1, column 0)
                Validation loadedValidation = loadedSheet.Validations.GetValidationInCell(1, 0);
                bool validationExists = loadedValidation != null;

                // -------------------- Output results --------------------
                Console.WriteLine("Comment verification:");
                Console.WriteLine($"  Exists: {commentExists}");
                Console.WriteLine($"  Note matches: {commentNote == "Original comment text"}");
                Console.WriteLine($"  Author matches: {commentAuthor == "Original Author"}");

                Console.WriteLine("Data validation verification:");
                Console.WriteLine($"  Exists: {validationExists}");
                if (validationExists)
                {
                    Console.WriteLine($"  Type matches: {loadedValidation.Type == ValidationType.WholeNumber}");
                    Console.WriteLine($"  Operator matches: {loadedValidation.Operator == OperatorType.Between}");
                    Console.WriteLine($"  Formula1 matches: {loadedValidation.Formula1 == "10"}");
                    Console.WriteLine($"  Formula2 matches: {loadedValidation.Formula2 == "100"}");
                }
            }
        }
    }
}
