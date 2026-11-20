// Title: Check if an Excel workbook is password‑protected using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells how to determine if an .xlsx file requires a password | detect password protection on an Excel workbook with Aspose.Cells .NET | verify Excel file encryption status before opening using Aspose.Cells in C# | Aspose.Cells load workbook without password and catch protection exception | method to test if an Excel file is password protected in a .NET application
// Tags: Aspose.Cells workbook password detection | C# check Excel file encryption status | load Excel workbook without password Aspose.Cells | CellsException handling for protected workbooks | determine .xlsx password protection .NET

using System;
using System.IO;
using Aspose.Cells;

// The IsPasswordProtected method validates the file path, attempts to load the workbook with Aspose.Cells' Workbook constructor without a password, and returns false when loading succeeds. If a CellsException containing the word "password" is thrown, the method returns true, indicating the workbook is password‑protected; all other exceptions are re‑thrown for the caller to handle.
public static class ExcelProtectionChecker
{
    /// <param name="filePath">Full path to the Excel file.</param>
    /// <returns>True if the file is password protected; otherwise, false.</returns>
    public static bool IsPasswordProtected(string filePath)
    {
        // Ensure the file exists before attempting to load it.
        if (!File.Exists(filePath))
            throw new FileNotFoundException($"The file \"{filePath}\" was not found.");

        try
        {
            // Attempt to load the workbook without a password.
            // If the file is not protected, this succeeds.
            var workbook = new Workbook(filePath);
            return false; // Loaded successfully, so not password protected.
        }
        catch (CellsException ex)
        {
            // Aspose.Cells throws CellsException when a password is required.
            // Check the exception message for password‑related text.
            if (!string.IsNullOrEmpty(ex.Message) &&
                ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return true; // Workbook is password protected.
            }

            // Re‑throw any other Aspose.Cells specific exceptions.
            throw;
        }
        catch (Exception)
        {
            // Re‑throw unexpected exceptions to be handled by the caller.
            throw;
        }
    }
}

public class Program
{
    // Entry point required for compilation.
    public static void Main(string[] args)
    {
        // Use the first argument as the file path, or a default placeholder.
        string filePath = args.Length > 0 ? args[0] : "example.xlsx";

        try
        {
            bool isProtected = ExcelProtectionChecker.IsPasswordProtected(filePath);
            Console.WriteLine($"File \"{filePath}\" password protected: {isProtected}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message.
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
