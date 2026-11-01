// Title: Unprotect an Excel worksheet in C# using Aspose.Cells with a password from an environment variable
// Common Searches: C# Aspose.Cells unprotect worksheet using password stored in environment variable | How to read worksheet password from environment variable and unprotect Excel file with Aspose | Aspose.Cells unprotect Excel sheet programmatically with .NET environment variable | Securely unprotect protected Excel worksheet in .NET without hard‑coding password | Retrieve SHEET_PASSWORD env variable and call Worksheet.Unprotect in Aspose.Cells
// Tags: Aspose.Cells worksheet unprotect with environment variable | C# load password from env variable for Excel protection | secure Excel sheet unprotection .NET | environment variable password handling Aspose.Cells | unprotect Excel workbook using Aspose.Cells API

using System;
using Aspose.Cells;

// The example reads the password from the SHEET_PASSWORD environment variable, loads an Excel workbook with Aspose.Cells, accesses the target worksheet, calls Worksheet.Unprotect with the retrieved password, and saves the updated file.
class Program
{
    static void Main()
    {
        // Retrieve the worksheet password from an environment variable
        string password = Environment.GetEnvironmentVariable("SHEET_PASSWORD");
        if (string.IsNullOrEmpty(password))
        {
            Console.WriteLine("Error: Environment variable 'SHEET_PASSWORD' is not set.");
            return;
        }

        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the worksheet you want to unprotect (e.g., the first sheet)
        Worksheet worksheet = workbook.Worksheets[0];

        // Unprotect the worksheet using the retrieved password
        worksheet.Unprotect(password);

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}
