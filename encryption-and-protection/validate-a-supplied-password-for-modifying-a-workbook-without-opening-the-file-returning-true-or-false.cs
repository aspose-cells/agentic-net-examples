// Title: Validate an Excel workbook's modify password without fully loading the file using Aspose.Cells for .NET
// Common Searches: how to test an Excel file's modify password with Aspose.Cells in C# | C# check if workbook password allows editing without opening the workbook | Aspose.Cells verify password for modifying XLSX without loading entire file | determine if Excel workbook is protected for editing using LoadOptions in .NET | validate workbook modify password programmatically with Aspose.Cells
// Tags: Aspose.Cells LoadOptions password verification | C# validate Excel workbook modify protection | catch CellsException for invalid workbook password | test workbook edit password without full load | use LoadOptions to check Excel file password

using System;
using System.IO;
using Aspose.Cells;

namespace Example
{
    // The ValidateModifyPassword method determines whether a supplied password can open an Excel workbook for modification. It first checks that the file exists, then creates LoadOptions with the password and attempts to load the workbook. If loading succeeds, the method returns true; if a CellsException or any other exception occurs, it returns false, indicating an invalid password or other load issue.
    public class WorkbookPasswordValidator
    {
        /// <param name="filePath">Full path to the workbook file.</param>
        /// <param name="password">Password to validate.</param>
        /// <returns>True if password is valid; false otherwise.</returns>
        public static bool ValidateModifyPassword(string filePath, string password)
        {
            // Ensure the file exists to avoid FileNotFoundException.
            if (!File.Exists(filePath))
                return false;

            // Prepare load options with the supplied password.
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = password
            };

            try
            {
                // Attempt to load the workbook using the password.
                Workbook workbook = new Workbook(filePath, loadOptions);
                // If loading succeeds, the password is valid.
                return true;
            }
            catch (CellsException)
            {
                // Invalid password or protection error.
                return false;
            }
            catch (Exception)
            {
                // Any other exception (e.g., corrupted file) treated as failure.
                return false;
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Example usage: provide path and password via command line or hard‑coded values.
            string path = args.Length > 0 ? args[0] : "sample.xlsx";
            string pwd = args.Length > 1 ? args[1] : "password";

            bool isValid = WorkbookPasswordValidator.ValidateModifyPassword(path, pwd);
            Console.WriteLine(isValid
                ? "Password is valid for modifying the workbook."
                : "Password is invalid or file could not be opened.");
        }
    }
}
