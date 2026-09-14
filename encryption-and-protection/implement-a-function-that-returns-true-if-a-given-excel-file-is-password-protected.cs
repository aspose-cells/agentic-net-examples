// Title: Check if an Excel workbook is password‑protected using Aspose.Cells for .NET (C#)
// AI Prompts: Write a C# method that takes a file path, attempts to open the workbook with Aspose.Cells without a password, and returns true when a CellsException indicates the file is password‑protected. | Create a utility that validates the existence of an Excel file, loads it via Aspose.Cells, and determines programmatically whether the workbook requires a password.
// Common Searches: how to programmatically detect password protection on an .xlsx file using Aspose.Cells C# | C# Aspose.Cells determine if a workbook requires a password before opening | catch CellsException to identify encrypted Excel workbook with Aspose.Cells | check if Excel file is encrypted without opening it using Aspose.Cells .NET | validate Excel file path and detect protection using Aspose.Cells
// Tags: Aspose.Cells detect password protection | C# check encrypted Excel workbook | load workbook without password Aspose.Cells | handle CellsException for protected files | validate Excel file path Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The IsPasswordProtected method verifies the provided file path, attempts to load the workbook with Aspose.Cells without supplying a password, returns false if loading succeeds, and returns true when a CellsException is caught, indicating the Excel file is password‑protected.
public static class ExcelProtectionHelper
{
    /// <param name="filePath">Full path to the Excel file.</param>
    /// <returns>True if the file requires a password; otherwise, false.</returns>
    public static bool IsPasswordProtected(string filePath)
    {
        if (string.IsNullOrWhiteSpace(filePath))
            throw new ArgumentException("File path must be provided.", nameof(filePath));

        if (!File.Exists(filePath))
            throw new FileNotFoundException($"The file '{filePath}' was not found.", filePath);

        try
        {
            // Attempt to load the workbook without a password.
            // If loading succeeds, the file is not password‑protected.
            Workbook workbook = new Workbook(filePath);
            return false;
        }
        catch (CellsException)
        {
            // Loading failed due to protection (or other Cells‑related issue).
            // Assume the workbook is password‑protected.
            return true;
        }
        catch (Exception)
        {
            // Re‑throw any other unexpected exceptions.
            throw;
        }
    }
}

// Optional entry point to satisfy console‑app projects.
public class Program
{
    public static void Main(string[] args)
    {
        // Example usage (adjust the path as needed).
        try
        {
            string path = args.Length > 0 ? args[0] : "sample.xlsx";

            // Ensure the example file exists before checking.
            if (!File.Exists(path))
            {
                Console.WriteLine($"File not found: {path}");
                return;
            }

            bool protectedFile = ExcelProtectionHelper.IsPasswordProtected(path);
            Console.WriteLine($"File '{path}' password protected: {protectedFile}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
