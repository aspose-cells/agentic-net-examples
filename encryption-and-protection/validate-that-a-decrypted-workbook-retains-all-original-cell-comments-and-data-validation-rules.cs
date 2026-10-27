// Title: How to confirm that cell comments and data validation rules survive workbook decryption with Aspose.Cells for .NET
// Common Searches: C# code to compare comments in original and decrypted Excel workbook using Aspose.Cells | verify that data validation objects remain unchanged after opening a password‑protected XLSX with Aspose.Cells | check integrity of cell comments after decrypting an encrypted workbook in .NET | Aspose.Cells compare worksheet validations between two workbooks | detect missing comments when loading an encrypted Excel file in C#
// Tags: compare worksheet comments Aspose.Cells | validate data validation preservation Aspose.Cells | load password protected XLSX Aspose.Cells | workbook integrity check .NET | cell comment dictionary Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace WorkbookIntegrityCheck
{
    // The example loads an original unencrypted workbook and an encrypted workbook (decrypted with a password) using Aspose.Cells for .NET, then verifies that both workbooks contain identical cell comments and data validation rules by building comment dictionaries and validation lists for each worksheet and comparing their counts and contents.
    class Program
    {
        static void Main()
        {
            // Paths to the original (unencrypted) and encrypted workbooks
            string originalPath = "original.xlsx";
            string encryptedPath = "encrypted.xlsx";

            // Password used to encrypt the workbook
            string password = "YourPassword";

            // Verify that input files exist
            if (!File.Exists(originalPath))
            {
                Console.WriteLine($"File not found: {originalPath}");
                return;
            }

            if (!File.Exists(encryptedPath))
            {
                Console.WriteLine($"File not found: {encryptedPath}");
                return;
            }

            try
            {
                // Load the original workbook (no password needed)
                Workbook wbOriginal = new Workbook(originalPath);

                // Load the encrypted workbook using the password (this decrypts it in memory)
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
                {
                    Password = password
                };
                Workbook wbDecrypted = new Workbook(encryptedPath, loadOptions);

                // Perform validation
                bool isValid = ValidateComments(wbOriginal, wbDecrypted) &&
                               ValidateDataValidations(wbOriginal, wbDecrypted);

                Console.WriteLine(isValid
                    ? "Decrypted workbook retains all original comments and data validation rules."
                    : "Validation failed: differences were found.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Compare comments in all worksheets
        private static bool ValidateComments(Workbook original, Workbook decrypted)
        {
            bool allMatch = true;

            // Ensure both workbooks have the same number of worksheets
            if (original.Worksheets.Count != decrypted.Worksheets.Count)
            {
                Console.WriteLine("Worksheet count mismatch.");
                return false;
            }

            for (int i = 0; i < original.Worksheets.Count; i++)
            {
                Worksheet wsOriginal = original.Worksheets[i];
                Worksheet wsDecrypted = decrypted.Worksheets[i];

                // Build dictionaries keyed by cell address (e.g., "A1")
                var originalComments = BuildCommentDictionary(wsOriginal);
                var decryptedComments = BuildCommentDictionary(wsDecrypted);

                // Compare counts
                if (originalComments.Count != decryptedComments.Count)
                {
                    Console.WriteLine($"Comment count mismatch in sheet '{wsOriginal.Name}'.");
                    allMatch = false;
                }

                // Compare each comment
                foreach (var kvp in originalComments)
                {
                    if (!decryptedComments.TryGetValue(kvp.Key, out string decryptedText))
                    {
                        Console.WriteLine($"Missing comment at {kvp.Key} in sheet '{wsOriginal.Name}'.");
                        allMatch = false;
                        continue;
                    }

                    if (!string.Equals(kvp.Value, decryptedText, StringComparison.Ordinal))
                    {
                        Console.WriteLine($"Comment text mismatch at {kvp.Key} in sheet '{wsOriginal.Name}'.");
                        allMatch = false;
                    }
                }
            }

            return allMatch;
        }

        // Helper to create a dictionary of comments for a worksheet
        private static Dictionary<string, string> BuildCommentDictionary(Worksheet sheet)
        {
            var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (Comment comment in sheet.Comments)
            {
                string address = CellsHelper.CellIndexToName(comment.Row, comment.Column);
                // Treat null notes as empty strings
                dict[address] = comment.Note ?? string.Empty;
            }
            return dict;
        }

        // Compare data validation rules in all worksheets
        private static bool ValidateDataValidations(Workbook original, Workbook decrypted)
        {
            bool allMatch = true;

            if (original.Worksheets.Count != decrypted.Worksheets.Count)
            {
                Console.WriteLine("Worksheet count mismatch.");
                return false;
            }

            for (int i = 0; i < original.Worksheets.Count; i++)
            {
                Worksheet wsOriginal = original.Worksheets[i];
                Worksheet wsDecrypted = decrypted.Worksheets[i];

                var originalValidations = BuildValidationList(wsOriginal);
                var decryptedValidations = BuildValidationList(wsDecrypted);

                if (originalValidations.Count != decryptedValidations.Count)
                {
                    Console.WriteLine($"Data validation count mismatch in sheet '{wsOriginal.Name}'.");
                    allMatch = false;
                }

                // Compare each original validation with a matching one in the decrypted workbook
                foreach (Validation valOrig in originalValidations)
                {
                    bool matchFound = false;

                    foreach (Validation valDec in decryptedValidations)
                    {
                        if (valOrig.Type == valDec.Type &&
                            valOrig.Operator == valDec.Operator &&
                            string.Equals(valOrig.Formula1, valDec.Formula1, StringComparison.Ordinal) &&
                            string.Equals(valOrig.Formula2, valDec.Formula2, StringComparison.Ordinal))
                        {
                            matchFound = true;
                            break;
                        }
                    }

                    if (!matchFound)
                    {
                        Console.WriteLine($"Missing or mismatched data validation (Type={valOrig.Type}) in sheet '{wsOriginal.Name}'.");
                        allMatch = false;
                    }
                }
            }

            return allMatch;
        }

        // Helper to collect all validations from a worksheet into a list
        private static List<Validation> BuildValidationList(Worksheet sheet)
        {
            var list = new List<Validation>();
            foreach (Validation val in sheet.Validations)
            {
                list.Add(val);
            }
            return list;
        }
    }
}
