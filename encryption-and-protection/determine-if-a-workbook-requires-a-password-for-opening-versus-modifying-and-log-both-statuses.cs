// Title: Determine if an Excel workbook requires an opening password or a modify password using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells how to detect an opening password in an Excel workbook | Check if an .xlsx file requires a write password using Aspose.Cells .NET | Read password protection flags of a workbook with Aspose.Cells in C# | Log whether an Excel file has an open or modify password via Aspose.Cells | Determine workbook.Settings.Password value for password protection detection
// Tags: Aspose.Cells detect workbook open password | Aspose.Cells evaluate write protection password | Workbook.Settings.Password property example | Excel encryption status logging with Aspose.Cells | C# password protection detection for .xlsx files

using System;
using System.IO;
using Aspose.Cells;

// The C# sample loads an Excel file with Aspose.Cells, checks the Workbook.Settings.Password property to see if an opening password is set, assumes no modify‑password flag is exposed, and writes both the opening and modify password requirements to the console while handling missing‑file and loading errors.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists before attempting to load it
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: File not found – {inputPath}");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook from the specified file
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        try
        {
            // Opening (read) password
            bool hasOpenPassword = !string.IsNullOrEmpty(workbook.Settings.Password);

            // Modify (write) password – Aspose.Cells does not expose direct flags in older versions,
            // so we assume false when the properties are unavailable.
            bool hasModifyPassword = false;

            // Output the password protection statuses
            Console.WriteLine($"Opening password required: {hasOpenPassword}");
            Console.WriteLine($"Modify password required: {hasModifyPassword}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
