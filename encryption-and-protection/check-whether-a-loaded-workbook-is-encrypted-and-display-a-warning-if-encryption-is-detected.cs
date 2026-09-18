// Title: Detect password protection on an Excel workbook using Aspose.Cells for .NET and output a warning
// AI Prompts: Generate C# code that loads an Excel file with Aspose.Cells, inspects the Settings.Password property to see if the workbook is password‑protected, and prints a warning if a password is detected. | Show how to open an Excel workbook with Aspose.Cells without supplying a password, handle any loading exceptions, and report the encryption status based on the workbook's protection settings.
// Common Searches: how to programmatically determine if an Excel file is password protected with Aspose.Cells in C# | Aspose.Cells .NET check workbook encryption status before opening | C# detect password‑protected workbook using Aspose.Cells Settings.Password | display warning when loading encrypted Excel workbook with Aspose.Cells | verify if Excel workbook requires a password using Aspose.Cells API
// Tags: Aspose.Cells detect workbook password protection | C# check Excel encryption Aspose.Cells | Aspose.Cells Settings.Password usage | C# load Excel without password handling | Aspose.Cells workbook protection status

using System;
using System.IO;
using Aspose.Cells;

// C# program that loads an Excel file with Aspose.Cells, examines the Settings.Password property to infer password protection, and writes a warning or confirmation to the console.
class Program
{
    static void Main()
    {
        // Path to the Excel file to be checked
        string filePath = "sample.xlsx";

        // Verify that the file exists before attempting to load it
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Error: The file \"{filePath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook (no password provided)
            Workbook workbook = new Workbook(filePath);

            // Aspose.Cells .NET does not expose a direct IsPasswordProtected property.
            // Attempt to access the workbook's protection settings to infer protection status.
            bool isProtected = workbook.Settings?.Password != null && workbook.Settings.Password.Length > 0;

            if (isProtected)
            {
                Console.WriteLine("Warning: The loaded workbook appears to be password‑protected.");
            }
            else
            {
                Console.WriteLine("The workbook is not password‑protected.");
            }
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors during loading or processing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
