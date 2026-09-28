// Title: Check an encrypted Excel workbook password with Aspose.Cells in C# without fully loading the file
// AI Prompts: Generate a C# method that receives a file path and a password, uses Aspose.Cells LoadOptions to attempt opening the workbook, and returns true only if the password is correct. | Write code that catches the specific CellsException for an incorrect password when loading a protected Excel file with Aspose.Cells, and returns a boolean indicating validation success.
// Common Searches: how to validate password of a password‑protected Excel file using Aspose.Cells C# | C# verify encrypted workbook password without reading its contents | Aspose.Cells lightweight password check for Excel workbook | detect wrong password when loading protected Excel file with Aspose.Cells | C# method to test Excel file password without opening the workbook
// Tags: Aspose.Cells LoadOptions password validation | C# encrypted Excel workbook password check | Excel password verification without full load | handle CellsException incorrect password | lightweight workbook password test Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace Example
{
    // Provides a static VerifyPassword method that creates LoadOptions with the supplied password, attempts to instantiate a Workbook, returns true on success, returns false when a CellsException indicates an incorrect password, and rethrows other exceptions as InvalidOperationException.
    public class WorkbookPasswordValidator
    {
        /// <param name="filePath">Full path to the encrypted Excel file.</param>
        /// <param name="password">Password to validate.</param>
        /// <returns>True if the password is correct; otherwise, false.</returns>
        public static bool VerifyPassword(string filePath, string password)
        {
            // Ensure the file exists before attempting to load it.
            if (!File.Exists(filePath))
                throw new FileNotFoundException("Workbook file not found.", filePath);

            // Prepare load options with the supplied password.
            var loadOptions = new LoadOptions(LoadFormat.Auto)
            {
                Password = password
            };

            try
            {
                // Attempt to load the workbook using the password.
                // If the password is wrong, Aspose.Cells throws a CellsException.
                using (var wb = new Workbook(filePath, loadOptions))
                {
                    // If we reach this point, the password is correct.
                    return true;
                }
            }
            catch (CellsException ex)
            {
                // Aspose.Cells throws a generic CellsException for incorrect passwords.
                // Detect it via the exception message.
                if (ex.Message != null && ex.Message.IndexOf("incorrect password", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return false; // Password did not match.
                }

                // Re‑throw for other CellsException scenarios (e.g., corrupted file).
                throw new InvalidOperationException("Failed to validate workbook password.", ex);
            }
            catch (Exception ex)
            {
                // Wrap any other unexpected exceptions.
                throw new InvalidOperationException("Failed to validate workbook password.", ex);
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Example usage: provide file path and password via command‑line arguments or hard‑code them.
            string filePath = args.Length > 0 ? args[0] : "encrypted.xlsx";
            string password = args.Length > 1 ? args[1] : "yourPassword";

            try
            {
                bool isValid = WorkbookPasswordValidator.VerifyPassword(filePath, password);
                Console.WriteLine(isValid ? "Password is correct." : "Password is incorrect.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
