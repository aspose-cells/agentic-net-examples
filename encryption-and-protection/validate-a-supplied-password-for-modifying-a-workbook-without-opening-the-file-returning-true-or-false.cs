// Title: Validate an Excel workbook's modify password without fully loading the file using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells verify modify password for an .xlsx file | how to test if an Excel workbook is editable with a password using Aspose.Cells | validate Excel file edit password without fully loading the workbook in .NET | check workbook modification rights programmatically with Aspose.Cells LoadOptions | determine if supplied password can open an encrypted Excel workbook for editing in C#
// Tags: Aspose.Cells LoadOptions password parameter | C# check Excel workbook edit capability | detect invalid workbook password via CellsException | programmatic Excel protection verification | validate workbook modification rights without full load

using System;
using System.IO;
using Aspose.Cells;

// The WorkbookPasswordValidator class provides a static method ValidateModifyPassword that accepts a file path and a password, attempts to open the workbook using Aspose.Cells LoadOptions with the Password property, returns true when the workbook loads successfully (indicating the password grants modification rights), returns false when a CellsException is caught (incorrect password), and rethrows any other unexpected exceptions.
public static class WorkbookPasswordValidator
{
    // Returns true if the supplied password can open the workbook for modification, false otherwise.
    public static bool ValidateModifyPassword(string filePath, string password)
    {
        if (!File.Exists(filePath))
            throw new FileNotFoundException($"File not found: {filePath}");

        try
        {
            // Attempt to load the workbook with the provided password.
            var loadOptions = new LoadOptions(LoadFormat.Xlsx) { Password = password };
            var workbook = new Workbook(filePath, loadOptions);
            return true; // Loaded successfully, password is correct.
        }
        catch (CellsException)
        {
            // Assume password is incorrect if a CellsException is thrown.
            return false;
        }
        catch
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
        // Expecting: <filePath> <password>
        if (args.Length < 2)
        {
            Console.WriteLine("Usage: <exe> <filePath> <password>");
            return;
        }

        string filePath = args[0];
        string password = args[1];

        try
        {
            bool canModify = WorkbookPasswordValidator.ValidateModifyPassword(filePath, password);
            Console.WriteLine(canModify ? "Password is correct." : "Password is incorrect.");
        }
        catch (FileNotFoundException fnfEx)
        {
            Console.WriteLine(fnfEx.Message);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
