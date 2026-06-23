using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    public class WorkbookPasswordChecker
    {
        // Returns true if the supplied password is required (i.e., matches) to modify the workbook.
        public static bool IsPasswordRequiredToModify(string filePath, string password)
        {
            try
            {
                // Prevent FileNotFoundException
                if (!File.Exists(filePath))
                    throw new FileNotFoundException($"Workbook file not found: {filePath}");

                // Load the workbook from the specified file.
                Workbook workbook = new Workbook(filePath);

                // If the workbook is not write‑protected, no password is required.
                if (!workbook.Settings.WriteProtection.IsWriteProtected)
                    return false;

                // Validate the provided password against the write‑protection password.
                return workbook.Settings.WriteProtection.ValidatePassword(password);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error checking password: {ex.Message}");
                return false;
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Example usage: file path and password can be passed as command‑line arguments.
            string filePath = args.Length > 0 ? args[0] : "sample.xlsx";
            string password = args.Length > 1 ? args[1] : string.Empty;

            bool requiresPassword = WorkbookPasswordChecker.IsPasswordRequiredToModify(filePath, password);
            Console.WriteLine($"Password required to modify: {requiresPassword}");
        }
    }
}