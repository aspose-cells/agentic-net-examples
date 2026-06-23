using System;
using System.IO;
using Aspose.Cells;

public class WorkbookPasswordValidator
{
    /// <summary>
    /// Validates whether the supplied password can modify (write‑protect) the workbook.
    /// The validation is performed directly on the file stream without fully loading the workbook.
    /// </summary>
    /// <param name="filePath">Full path to the Excel file.</param>
    /// <param name="password">Password to validate.</param>
    /// <returns>True if the password is correct; otherwise false.</returns>
    public static bool ValidateModifyPassword(string filePath, string password)
    {
        // Open the workbook file as a read‑only stream.
        using (FileStream stream = File.OpenRead(filePath))
        {
            // FileFormatUtil.VerifyPassword checks the password for encrypted OOXML files.
            // It returns true when the password matches the one used to protect the file.
            return FileFormatUtil.VerifyPassword(stream, password);
        }
    }

    // Example usage
    public static void Main()
    {
        string path = "ProtectedWorkbook.xlsx";
        string correctPassword = "owner";
        string wrongPassword = "1234";

        bool isCorrect = ValidateModifyPassword(path, correctPassword);
        Console.WriteLine($"Password '{correctPassword}' is valid: {isCorrect}");

        bool isWrong = ValidateModifyPassword(path, wrongPassword);
        Console.WriteLine($"Password '{wrongPassword}' is valid: {isWrong}");
    }
}