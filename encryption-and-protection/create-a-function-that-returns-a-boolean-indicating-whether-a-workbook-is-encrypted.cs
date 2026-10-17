// Title: How to determine if an Excel workbook is encrypted using Aspose.Cells for .NET in C#
// Common Searches: C# code to verify whether an Excel workbook is password protected with Aspose.Cells | detect encrypted Excel file without supplying a password using Aspose.Cells LoadOptions | Aspose.Cells IsEncrypted property fallback example in .NET | catch Aspose.Cells exception to identify required workbook password in C# | method to return true if a workbook requires a password when loading with Aspose.Cells
// Tags: Aspose.Cells detect encrypted workbook | C# check Excel file encryption Aspose.Cells | LoadOptions Auto load workbook without password | Workbook IsEncrypted reflection C# | exception handling password required Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The IsWorkbookEncrypted method takes a file path, verifies the file exists, and tries to load the workbook with LoadOptions(LoadFormat.Auto). It first attempts to read the Workbook.IsEncrypted property via reflection; if the property is present its boolean value is returned. If the property is missing, the method infers encryption by catching exceptions whose messages contain “password” or “encrypted”. Unrelated exceptions are re‑thrown. The result is true when the workbook is password‑protected, otherwise false.
public static class WorkbookHelper
{
    /// <param name="filePath">Full path to the workbook file.</param>
    /// <returns>True if the workbook is encrypted; otherwise, false.</returns>
    public static bool IsWorkbookEncrypted(string filePath)
    {
        // Ensure the file exists to avoid FileNotFoundException.
        if (!File.Exists(filePath))
            throw new FileNotFoundException($"The file \"{filePath}\" was not found.");

        try
        {
            // Load the workbook without a password.
            var loadOptions = new LoadOptions(LoadFormat.Auto);
            var workbook = new Workbook(filePath, loadOptions);

            // Some Aspose.Cells versions expose IsEncrypted property.
            var isEncryptedProp = typeof(Workbook).GetProperty("IsEncrypted");
            if (isEncryptedProp != null && isEncryptedProp.PropertyType == typeof(bool))
            {
                return (bool)isEncryptedProp.GetValue(workbook);
            }

            // If the property is unavailable, assume the file is not encrypted
            // because loading succeeded without a password.
            return false;
        }
        catch (Exception ex)
        {
            // Aspose.Cells throws an exception when a password is required.
            // Detect this by checking the exception message for keywords.
            if (ex.Message != null &&
                (ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0 ||
                 ex.Message.IndexOf("encrypted", StringComparison.OrdinalIgnoreCase) >= 0))
            {
                return true;
            }

            // Re‑throw any other unexpected exceptions.
            throw;
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        try
        {
            // Example usage: provide the workbook path as the first argument.
            if (args.Length == 0)
            {
                Console.WriteLine("Please provide the full path to the workbook file as an argument.");
                return;
            }

            string filePath = args[0];
            bool encrypted = WorkbookHelper.IsWorkbookEncrypted(filePath);
            Console.WriteLine($"Workbook \"{filePath}\" encrypted: {encrypted}");
        }
        catch (FileNotFoundException fnfEx)
        {
            Console.WriteLine(fnfEx.Message);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
