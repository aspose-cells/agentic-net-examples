// Title: Validate an Excel workbook password without fully loading the file using Aspose.Cells for .NET
// AI Prompts: Generate a C# method that uses Aspose.Cells LoadOptions with a supplied password to attempt opening an .xlsx file and returns true only when the password unlocks the workbook. | Write code that loads a protected workbook via Aspose.Cells, catches CellsException, and provides a boolean indicating whether the given password is correct.
// Common Searches: how to check if an Excel .xlsx password is correct in C# using Aspose.Cells without opening the workbook | Aspose.Cells C# verify workbook protection programmatically | C# method to test Excel file password validity with LoadOptions | determine if Excel file is password‑protected using Aspose.Cells .NET | validate protected workbook password without reading its content in C#
// Tags: Aspose.Cells LoadOptions password verification | C# validate protected .xlsx workbook | Excel password check using CellsException | verify workbook encryption without full load | password validation for Aspose.Cells protected files

using System;
using System.IO;
using Aspose.Cells;

// Provides a static ValidatePassword method that creates LoadOptions with the given password, attempts to load the workbook, and returns true if loading succeeds; false is returned for invalid passwords, missing files, or other errors, handling CellsException appropriately.
public class WorkbookPasswordValidator
{
    /// <param name="filePath">Full path to the workbook file.</param>
    /// <param name="password">Password to validate.</param>
    /// <returns>True if the password is correct; otherwise false.</returns>
    public static bool ValidatePassword(string filePath, string password)
    {
        // Ensure the file exists before attempting to load.
        if (!File.Exists(filePath))
        {
            return false;
        }

        try
        {
            // Set load options with the supplied password.
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = password
            };

            // Attempt to load the workbook using the password.
            // If the password is incorrect, Aspose.Cells throws a CellsException.
            Workbook workbook = new Workbook(filePath, loadOptions);

            // Loading succeeded – the password is valid.
            return true;
        }
        catch (CellsException)
        {
            // Loading failed due to an invalid password.
            return false;
        }
        catch (Exception)
        {
            // Any other exception (e.g., file not found) is treated as a failure.
            return false;
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        // Expecting two arguments: file path and password.
        if (args.Length < 2)
        {
            Console.WriteLine("Usage: <filePath> <password>");
            return;
        }

        string filePath = args[0];
        string password = args[1];

        bool isValid = WorkbookPasswordValidator.ValidatePassword(filePath, password);
        Console.WriteLine($"Password valid: {isValid}");
    }
}
