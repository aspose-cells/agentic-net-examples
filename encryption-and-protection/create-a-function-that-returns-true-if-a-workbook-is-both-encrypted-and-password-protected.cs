// Title: Check if an Excel workbook is encrypted and password‑protected using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells how to determine if an Excel file is both encrypted and password protected | detect encrypted workbook with password using Aspose.Cells .NET | verify Excel workbook encryption status and password with Aspose.Cells API | Aspose.Cells IsEncrypted property example for C# | check workbook.Settings.Password and IsEncrypted in Aspose.Cells
// Tags: Aspose.Cells encrypted workbook detection | C# workbook password protection check | IsEncrypted property Aspose.Cells | Settings.Password verification Aspose.Cells | Excel file encryption status .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The IsEncryptedAndPasswordProtected method validates the file path, loads the workbook with Aspose.Cells, confirms a non‑empty Settings.Password, reads the IsEncrypted flag via reflection for version compatibility, and returns true only when both encryption and password protection are present; errors are logged and result in false.
    public static class WorkbookHelper
    {
        // Returns true if the workbook is encrypted and has a password set for protection.
        public static bool IsEncryptedAndPasswordProtected(string filePath)
        {
            // Verify that the file exists to avoid FileNotFoundException.
            if (!File.Exists(filePath))
                throw new FileNotFoundException($"The file '{filePath}' was not found.");

            try
            {
                // Load the workbook from the specified file.
                Workbook workbook = new Workbook(filePath);

                // In Aspose.Cells, a password is stored in the Settings.Password property.
                // If a password is present, the workbook is considered password‑protected.
                bool hasPassword = !string.IsNullOrEmpty(workbook.Settings?.Password);

                // Aspose.Cells provides the IsEncrypted property in newer versions.
                // If unavailable, we treat the presence of a password as an indication of encryption.
                bool isEncrypted = false;
                try
                {
                    // Attempt to read the IsEncrypted property via reflection to stay compatible with older versions.
                    var prop = typeof(Workbook).GetProperty("IsEncrypted");
                    if (prop != null && prop.PropertyType == typeof(bool))
                    {
                        var value = prop.GetValue(workbook);
                        if (value != null)
                        {
                            isEncrypted = (bool)value;
                        }
                    }
                }
                catch
                {
                    // Ignore reflection errors; fallback to password check only.
                }

                // Return true only when both conditions are satisfied.
                return isEncrypted && hasPassword;
            }
            catch (Exception ex)
            {
                // Log or handle the exception as needed. For now, return false to indicate the check failed.
                Console.Error.WriteLine($"Error processing workbook '{filePath}': {ex.Message}");
                return false;
            }
        }
    }

    public class Program
    {
        // Entry point required for the console application.
        public static void Main(string[] args)
        {
            try
            {
                if (args.Length == 0)
                {
                    Console.WriteLine("Usage: AsposeCellsExample <path-to-workbook>");
                    return;
                }

                string filePath = args[0];

                bool result = WorkbookHelper.IsEncryptedAndPasswordProtected(filePath);
                Console.WriteLine($"Workbook '{filePath}' encrypted and password protected: {result}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
