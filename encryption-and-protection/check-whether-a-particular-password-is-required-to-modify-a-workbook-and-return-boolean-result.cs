// Title: Determine whether an Excel workbook requires a password and validate it with Aspose.Cells for .NET
// AI Prompts: Write a C# method that tries to open an Excel file with Aspose.Cells and returns true only when the workbook is encrypted and the supplied password successfully loads it. | Build a console application that accepts a file path and optional password, uses Aspose.Cells LoadOptions to detect encryption, and prints a boolean indicating if the password is correct.
// Common Searches: how to programmatically check if an Excel file is password protected using Aspose.Cells in C# | C# Aspose.Cells load encrypted workbook with password and verify correctness | detect workbook encryption status before opening with Aspose.Cells .NET | validate Excel file password without opening the file using Aspose.Cells | Aspose.Cells determine if password is required for a given Excel workbook
// Tags: Aspose.Cells workbook password verification | C# detect encrypted Excel workbook | LoadOptions password validation Aspose.Cells | check Excel file protection status .NET | programmatic Excel encryption detection Aspose

using System;
using System.IO;
using Aspose.Cells;

// The example provides a WorkbookPasswordChecker.IsPasswordRequired method that first attempts to load an Excel file without a password using Aspose.Cells; if that fails, it retries with the supplied password via LoadOptions. It returns true only when the workbook is encrypted and the password is correct, otherwise false.
public class WorkbookPasswordChecker
{
    /// <param name="filePath">Full path to the Excel file.</param>
    /// <param name="password">Password to test.</param>
    /// <returns>
    /// True if the file is password‑protected and the supplied password is correct;
    /// false if the file is not protected or the password is incorrect.
    /// </returns>
    public static bool IsPasswordRequired(string filePath, string password)
    {
        // Ensure the file exists before attempting to load it.
        if (!File.Exists(filePath))
            throw new FileNotFoundException("The specified Excel file was not found.", filePath);

        // First, try to open the workbook without a password.
        try
        {
            Workbook wb = new Workbook(filePath);
            // Loaded successfully – no password is required.
            return false;
        }
        catch (CellsException)
        {
            // An exception likely means the workbook is encrypted; try with the supplied password.
            try
            {
                LoadOptions loadOptions = new LoadOptions();
                loadOptions.Password = password;
                Workbook wb = new Workbook(filePath, loadOptions);
                // Loaded successfully with the password – password is required and correct.
                return true;
            }
            catch (CellsException)
            {
                // Loading failed even with the password – password is incorrect.
                return false;
            }
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        // Example usage:
        // args[0] = path to Excel file, args[1] = password to test (optional)
        if (args.Length == 0)
        {
            Console.WriteLine("Please provide the path to the Excel file as the first argument.");
            return;
        }

        string filePath = args[0];
        string password = args.Length > 1 ? args[1] : string.Empty;

        try
        {
            bool result = WorkbookPasswordChecker.IsPasswordRequired(filePath, password);
            Console.WriteLine(result
                ? "The workbook is password‑protected and the supplied password is correct."
                : "The workbook is either not password‑protected or the supplied password is incorrect.");
        }
        catch (FileNotFoundException ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
