// Title: Check whether an Excel workbook is encrypted using Aspose.Cells in C#
// AI Prompts: Write a C# method that loads an .xlsx file with Aspose.Cells LoadOptions and returns true only when a CellsException indicates the workbook is encrypted. | Generate code that validates the file path, attempts to open the workbook without a password, and uses exception handling to determine encryption status.
// Common Searches: how to detect password protected Excel file with Aspose.Cells in C# | C# Aspose.Cells check if workbook is encrypted before opening | determine if .xlsx is encrypted using LoadOptions Aspose.Cells | catch CellsException to identify encrypted workbook in .NET
// Tags: detect encrypted workbook Aspose.Cells | check Excel file password protection C# | load encrypted .xlsx with LoadOptions Aspose | handle CellsException encryption detection | verify workbook encryption status .NET

using System;
using System.IO;
using Aspose.Cells;

// Provides a static C# helper method that verifies a file path, attempts to load the workbook with Aspose.Cells LoadOptions, and returns false if loading succeeds. If a CellsException is thrown, the method returns true, indicating the workbook is encrypted or password‑protected.
public static class WorkbookHelper
{
    /// <param name="filePath">Full path to the workbook file.</param>
    /// <returns>True if the workbook is encrypted; otherwise, false.</returns>
    public static bool IsWorkbookEncrypted(string filePath)
    {
        // Ensure the file exists before attempting to load it.
        if (!File.Exists(filePath))
            throw new FileNotFoundException($"File not found: {filePath}");

        try
        {
            // Attempt to load the workbook without a password.
            // If the file is encrypted, Aspose.Cells will throw a CellsException.
            var loadOptions = new LoadOptions(LoadFormat.Xlsx);
            var workbook = new Workbook(filePath, loadOptions);
            // Loaded successfully, therefore not encrypted.
            return false;
        }
        catch (CellsException)
        {
            // The workbook is likely encrypted (or password‑protected).
            return true;
        }
        catch (Exception)
        {
            // Re‑throw any other unexpected exceptions.
            throw;
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        // Example usage: check if a workbook is encrypted.
        if (args.Length == 0)
        {
            Console.WriteLine("Please provide the path to the workbook file as an argument.");
            return;
        }

        string filePath = args[0];

        try
        {
            bool isEncrypted = WorkbookHelper.IsWorkbookEncrypted(filePath);
            Console.WriteLine($"Workbook encrypted: {isEncrypted}");
        }
        catch (FileNotFoundException ex)
        {
            Console.WriteLine(ex.Message);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
