// Title: Determine whether an Excel workbook is encrypted and password‑protected using Aspose.Cells for .NET
// Common Searches: how to check if an .xlsx file is password protected with Aspose.Cells in C# | detect encrypted Excel workbook using Aspose.Cells LoadOptions without providing a password | C# method to verify workbook encryption status with Aspose.Cells | catch CellsException for password‑required Excel files using Aspose.Cells | validate workbook protection before opening it with Aspose.Cells .NET
// Tags: Aspose.Cells encrypted workbook detection | C# Excel password protection verification | LoadOptions password‑less loading Aspose.Cells | CellsException handling for protected workbooks | verify workbook encryption status .NET

using System;
using System.IO;
using Aspose.Cells;

// The IsWorkbookEncryptedAndPasswordProtected method checks the file path, attempts to load the workbook with default LoadOptions, and returns true when a CellsException containing the word "password" is thrown, indicating the workbook is both encrypted and password‑protected; otherwise it returns false and propagates unrelated exceptions.
public static class WorkbookHelper
{
    /// <param name="filePath">Full path to the workbook file.</param>
    /// <returns>True if the workbook is encrypted and password‑protected; otherwise, false.</returns>
    public static bool IsWorkbookEncryptedAndPasswordProtected(string filePath)
    {
        // Verify that the file exists to avoid FileNotFoundException.
        if (!File.Exists(filePath))
            throw new FileNotFoundException($"The workbook file was not found: {filePath}");

        try
        {
            // Attempt to load the workbook without a password.
            var loadOptions = new LoadOptions(); // No password supplied
            var workbook = new Workbook(filePath, loadOptions);

            // Loading succeeded, so the workbook is not password protected.
            return false;
        }
        // Aspose.Cells throws CellsException when a password is required.
        catch (CellsException ex)
        {
            // Check if the exception message indicates a password issue.
            if (ex.Message != null &&
                ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return true;
            }

            // Re‑throw if it's a different CellsException.
            throw;
        }
        // Re‑throw any other unexpected exceptions.
        catch (Exception)
        {
            throw;
        }
    }
}

// Simple entry point for testing the helper method.
public class Program
{
    public static void Main(string[] args)
    {
        // Example usage: replace with an actual file path when testing.
        string workbookPath = @"C:\Temp\Sample.xlsx";

        try
        {
            bool isProtected = WorkbookHelper.IsWorkbookEncryptedAndPasswordProtected(workbookPath);
            Console.WriteLine(isProtected
                ? "The workbook is encrypted and password protected."
                : "The workbook is not password protected.");
        }
        catch (FileNotFoundException fnfEx)
        {
            Console.WriteLine(fnfEx.Message);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
