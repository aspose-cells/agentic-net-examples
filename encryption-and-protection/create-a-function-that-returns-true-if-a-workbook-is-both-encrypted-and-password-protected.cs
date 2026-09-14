// Title: Determine whether an Excel workbook is encrypted and password protected with Aspose.Cells in C#
// AI Prompts: Write a C# function that opens an Excel file using Aspose.Cells with no password supplied and returns true only if workbook.Settings.IsEncrypted and workbook.Settings.IsProtected are both true. | Add robust error handling so the function throws FileNotFoundException for a missing path and returns false when the workbook cannot be loaded because a password is required. | Create a console application that accepts a file path argument, calls the encryption‑and‑protection check method, and prints the result to the console.
// Common Searches: aspocells how to check if an Excel file is both encrypted and password protected in C# | c# detect encrypted workbook using Aspose.Cells LoadOptions without password | verify Excel workbook encryption and protection status with Aspose.Cells API | Aspose.Cells IsEncrypted and IsProtected properties example | C# program to determine if Excel file requires a password using Aspose.Cells
// Tags: Aspose.Cells workbook encryption detection | Aspose.Cells password protection check | LoadOptions empty password C# | Workbook.Settings IsEncrypted IsProtected | C# verify Excel file encryption status

using System;
using System.IO;
using Aspose.Cells;

// Provides a C# helper method that loads an Excel workbook with Aspose.Cells using an empty password, then returns true only when both workbook.Settings.IsEncrypted and workbook.Settings.IsProtected are set, with proper file‑existence checks and graceful handling of load failures.
public static class WorkbookHelper
{
    /// <param name="filePath">Full path to the workbook file.</param>
    /// <returns>True when both encryption and password protection are present; otherwise false.</returns>
    public static bool IsWorkbookEncryptedAndPasswordProtected(string filePath)
    {
        // Ensure the file exists to prevent FileNotFoundException.
        if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
        {
            throw new FileNotFoundException($"Workbook file not found: {filePath}");
        }

        try
        {
            // Load the workbook without providing a password.
            var loadOptions = new LoadOptions
            {
                Password = string.Empty // No password supplied.
            };

            var workbook = new Workbook(filePath, loadOptions);

            // Check encryption and password protection flags.
            return workbook.Settings.IsEncrypted && workbook.Settings.IsProtected;
        }
        catch (Exception)
        {
            // If loading fails (e.g., due to missing password), return false.
            return false;
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        // Simple demonstration of the helper method.
        if (args.Length == 0)
        {
            Console.WriteLine("Please provide the full path to a workbook file as an argument.");
            return;
        }

        string filePath = args[0];

        try
        {
            bool isEncryptedAndProtected = WorkbookHelper.IsWorkbookEncryptedAndPasswordProtected(filePath);
            Console.WriteLine($"Workbook encrypted and password protected: {isEncryptedAndProtected}");
        }
        catch (FileNotFoundException ex)
        {
            Console.WriteLine(ex.Message);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
