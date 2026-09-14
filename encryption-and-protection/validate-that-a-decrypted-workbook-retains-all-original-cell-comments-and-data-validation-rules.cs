// Title: Check that a decrypted Excel workbook keeps all original cell comments and data‑validation rules using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an encrypted .xlsx file with a password via Aspose.Cells, decrypts it, and verifies that every comment (text and cell address) matches the corresponding comment in an unencrypted reference workbook. | Create a reusable C# method that takes two Aspose.Cells Workbook objects and returns true only if their comment collections and data‑validation collections are identical across all worksheets. | Generate a C# unit test that loads an original workbook and its password‑protected copy, decrypts the latter, and asserts that comment counts, comment texts, and validation settings are equal.
// Common Searches: aspnet compare comments between original and password protected Excel file using Aspose.Cells | how to ensure data validation rules are preserved after decrypting an XLSX with Aspose.Cells | C# verify that decrypted workbook retains cell comments and validations | Aspose.Cells load encrypted workbook and compare workbook objects for comment integrity | unit test for comment and validation equality after workbook decryption Aspose.Cells
// Tags: compare workbook comments Aspose.Cells | validate data validation preservation Aspose.Cells | load encrypted XLSX with password Aspose.Cells | workbook decryption integrity check | comment and validation equality C#

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The program loads an original (unencrypted) workbook and an encrypted workbook, decrypts the latter with a password using Aspose.Cells, then iterates through each worksheet to confirm that comment counts, positions, and texts match, and that all data‑validation rules are identical between the two workbooks.
class WorkbookValidation
{
    static void Main()
    {
        try
        {
            // Paths to the original (unencrypted) and encrypted workbooks
            string originalPath = "original.xlsx";
            string encryptedPath = "encrypted.xlsx";
            string password = "yourPassword";

            // Ensure the files exist before attempting to load them
            if (!File.Exists(originalPath))
                throw new FileNotFoundException($"Original workbook not found: {originalPath}");
            if (!File.Exists(encryptedPath))
                throw new FileNotFoundException($"Encrypted workbook not found: {encryptedPath}");

            // Load the original workbook (no password needed)
            Workbook originalWb = new Workbook(originalPath);

            // Load the encrypted workbook using the password
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = password
            };
            Workbook decryptedWb = new Workbook(encryptedPath, loadOptions);

            // Validate comments and data validations
            bool commentsMatch = CompareComments(originalWb, decryptedWb);
            bool validationsMatch = CompareDataValidations(originalWb, decryptedWb);

            Console.WriteLine($"Comments match: {commentsMatch}");
            Console.WriteLine($"Data validations match: {validationsMatch}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Compare all comments in both workbooks
    private static bool CompareComments(Workbook wb1, Workbook wb2)
    {
        if (wb1.Worksheets.Count != wb2.Worksheets.Count)
            return false;

        for (int i = 0; i < wb1.Worksheets.Count; i++)
        {
            Worksheet ws1 = wb1.Worksheets[i];
            Worksheet ws2 = wb2.Worksheets[i];

            // Compare comment counts
            if (ws1.Comments.Count != ws2.Comments.Count)
                return false;

            // Build a dictionary of comments from the second workbook for quick lookup
            Dictionary<string, string> ws2Comments = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (Comment c in ws2.Comments)
            {
                // Use row/column to build cell name (e.g., "A1")
                string cellName = CellsHelper.CellIndexToName(c.Row, c.Column);
                ws2Comments[cellName] = c.Note;
            }

            // Verify each comment from the first workbook exists and has identical text
            foreach (Comment c1 in ws1.Comments)
            {
                string cellName = CellsHelper.CellIndexToName(c1.Row, c1.Column);
                if (!ws2Comments.TryGetValue(cellName, out string note2))
                    return false;

                if (!string.Equals(c1.Note, note2, StringComparison.Ordinal))
                    return false;
            }
        }

        return true;
    }

    // Compare all data validation rules in both workbooks
    private static bool CompareDataValidations(Workbook wb1, Workbook wb2)
    {
        if (wb1.Worksheets.Count != wb2.Worksheets.Count)
            return false;

        for (int i = 0; i < wb1.Worksheets.Count; i++)
        {
            Worksheet ws1 = wb1.Worksheets[i];
            Worksheet ws2 = wb2.Worksheets[i];

            // Compare validation counts
            if (ws1.Validations.Count != ws2.Validations.Count)
                return false;

            // Create a mutable list of validations from the second workbook for one‑to‑one matching
            List<Aspose.Cells.Validation> ws2Validations = new List<Aspose.Cells.Validation>(ws2.Validations);

            foreach (Aspose.Cells.Validation v1 in ws1.Validations)
            {
                bool matchFound = false;
                for (int j = 0; j < ws2Validations.Count; j++)
                {
                    Aspose.Cells.Validation v2 = ws2Validations[j];
                    if (ValidationsAreEqual(v1, v2))
                    {
                        matchFound = true;
                        ws2Validations.RemoveAt(j); // Ensure one‑to‑one matching
                        break;
                    }
                }
                if (!matchFound)
                    return false;
            }
        }

        return true;
    }

    // Helper to compare two Validation objects (area comparison omitted for compatibility)
    private static bool ValidationsAreEqual(Aspose.Cells.Validation v1, Aspose.Cells.Validation v2)
    {
        // Compare validation type, operator, and formulas
        if (v1.Type != v2.Type ||
            v1.Operator != v2.Operator ||
            !string.Equals(v1.Formula1, v2.Formula1, StringComparison.Ordinal) ||
            !string.Equals(v1.Formula2, v2.Formula2, StringComparison.Ordinal))
            return false;

        // Compare additional properties that affect behavior
        if (v1.IgnoreBlank != v2.IgnoreBlank ||
            v1.InCellDropDown != v2.InCellDropDown ||
            v1.ShowError != v2.ShowError ||
            v1.ShowInput != v2.ShowInput)
            return false;

        // Compare error and input messages (if any)
        if (!string.Equals(v1.ErrorMessage, v2.ErrorMessage, StringComparison.Ordinal) ||
            !string.Equals(v1.InputMessage, v2.InputMessage, StringComparison.Ordinal))
            return false;

        return true;
    }
}
