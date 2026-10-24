// Title: Protect an Excel worksheet with a password and allow cell formatting using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells protect worksheet with password and allow cell formatting | How to enable AllowFormattingCell while protecting an Excel sheet using Aspose.Cells | Set password protection on Excel worksheet but keep formatting options enabled in .NET | Aspose.Cells protect sheet and still let users change cell styles | Example code for protecting Excel worksheet with password and formatting permission in C#
// Tags: worksheet protection password Aspose.Cells | allow formatting cells on protected sheet C# | AllowFormattingCell property usage Aspose.Cells | protect Excel worksheet while permitting formatting | save protected workbook Aspose.Cells .NET

using Aspose.Cells;
using System;
using System.IO;

// Creates a new Workbook, enables the AllowFormattingCell option on the first worksheet, assigns a password to the sheet's Protection, ensures the output directory exists, and saves the workbook as a password‑protected Excel file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Allow users to format cells while protecting the sheet
            sheet.Protection.AllowFormattingCell = true;

            // Protect the worksheet with a password (use the Password property)
            sheet.Protection.Password = "MySecretPassword";

            // Define output file path
            string outputPath = "ProtectedWorkbook.xlsx";

            // Ensure the directory exists (handle possible null from GetDirectoryName)
            string fullPath = Path.GetFullPath(outputPath);
            string? directory = Path.GetDirectoryName(fullPath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
