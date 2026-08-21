// Title: C# – Validate an Excel Workbook’s Write‑Protection Password via Stream Using Aspose.Cells
// Description: Shows how to open an Excel file as a read‑only stream and call Aspose.Cells.FileFormatUtil.VerifyPassword to determine whether a supplied password matches the workbook’s write‑protection password, returning true or false without fully loading the workbook.
// Keywords: Aspose.Cells | C# | .NET | Excel password verification | write‑protected workbook | FileFormatUtil.VerifyPassword | validate workbook password | stream‑based password check | no full load | Excel security | global
// Common Searches: Aspose.Cells verify write‑protected Excel password without opening file | C# check Excel workbook password from stream | FileFormatUtil.VerifyPassword example | how to validate Excel file protection programmatically | validate Excel workbook password .NET
// Developer Intent: Check if a given password unlocks the write‑protection of an Excel workbook without loading the entire file into memory.
// Use Cases: Pre‑validate passwords in batch jobs before attempting to modify protected workbooks. | Expose an API that confirms edit rights for a supplied Excel file and password. | Filter a repository of Excel files to identify those that can be edited with a known password.
// AI Prompts: Generate C# code that uses Aspose.Cells to verify a workbook’s write‑protection password from a file stream and includes error handling. | Explain the difference in behavior of FileFormatUtil.VerifyPassword for encrypted versus write‑protected Excel files. | Provide a sample service that iterates over multiple Excel files, validates each password, and logs the results.

using System;
using System.IO;
using Aspose.Cells;

// Shows how to open an Excel file as a read‑only stream and call Aspose.Cells.FileFormatUtil.VerifyPassword to determine whether a supplied password matches the workbook’s write‑protection password, returning true or false without fully loading the workbook.
public class WorkbookPasswordValidator
{
    // Validates the supplied password for a workbook without fully loading the file.
    // Returns true if the password matches the write‑protection password, otherwise false.
    public static bool ValidatePassword(string filePath, string password)
    {
        // Open the workbook file as a read‑only stream.
        using (Stream stream = File.OpenRead(filePath))
        {
            // FileFormatUtil.VerifyPassword checks the password for encrypted or write‑protected workbooks
            // directly from the stream, avoiding full workbook loading.
            return FileFormatUtil.VerifyPassword(stream, password);
        }
    }

    // Demonstration of the validation method.
    public static void Main()
    {
        string workbookPath = "WriteProtectedWorkbook.xlsx";
        string passwordToTest = "owner";

        bool isPasswordValid = ValidatePassword(workbookPath, passwordToTest);
        Console.WriteLine($"Password '{passwordToTest}' is valid: {isPasswordValid}");
    }
}
