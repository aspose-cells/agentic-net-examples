using System;
using Aspose.Cells;

namespace AsposeCellsValidationDemo
{
    class Program
    {
        static void Main()
        {
            // -------------------- Create original workbook --------------------
            Workbook originalWb = new Workbook();
            Worksheet sheet = originalWb.Worksheets[0];

            // Add a comment to cell A1
            int commentIdx = sheet.Comments.Add("A1");
            sheet.Comments[commentIdx].Note = "Original comment text";
            sheet.Comments[commentIdx].Author = "OriginalAuthor";

            // Add data validation to cell B2 (whole number between 1 and 10)
            Validation validation = sheet.Validations[sheet.Validations.Add()];
            validation.Type = ValidationType.WholeNumber;
            validation.Operator = OperatorType.Between;
            validation.Formula1 = "1";
            validation.Formula2 = "10";

            CellArea area = new CellArea
            {
                StartRow = 1,   // B2 row index (0‑based)
                StartColumn = 1,
                EndRow = 1,
                EndColumn = 1
            };
            validation.AddArea(area);

            // Protect the workbook with a password (encryption)
            originalWb.Settings.Password = "SecretPwd";

            // Save the encrypted workbook
            string encryptedPath = "EncryptedWorkbook.xlsx";
            originalWb.Save(encryptedPath);
            originalWb.Dispose();

            // -------------------- Load (decrypt) workbook --------------------
            LoadOptions loadOpts = new LoadOptions
            {
                Password = "SecretPwd"
            };
            Workbook loadedWb = new Workbook(encryptedPath, loadOpts);
            Worksheet loadedSheet = loadedWb.Worksheets[0];

            // Verify comment
            bool commentExists = loadedSheet.Comments.Count > 0;
            string commentNote = commentExists ? loadedSheet.Comments[0].Note : string.Empty;
            string commentAuthor = commentExists ? loadedSheet.Comments[0].Author : string.Empty;

            // Verify data validation
            bool validationExists = loadedSheet.Validations.Count > 0;
            Validation loadedValidation = validationExists ? loadedSheet.Validations[0] : null;
            string validationInfo = validationExists
                ? $"Type={loadedValidation.Type}, Operator={loadedValidation.Operator}, Formula1={loadedValidation.Formula1}, Formula2={loadedValidation.Formula2}"
                : "None";

            // Output verification results
            Console.WriteLine("Comment exists: " + commentExists);
            Console.WriteLine("Comment note matches: " + (commentNote == "Original comment text"));
            Console.WriteLine("Comment author matches: " + (commentAuthor == "OriginalAuthor"));
            Console.WriteLine("Validation exists: " + validationExists);
            Console.WriteLine("Validation details: " + validationInfo);

            // Clean up
            loadedWb.Dispose();
        }
    }
}