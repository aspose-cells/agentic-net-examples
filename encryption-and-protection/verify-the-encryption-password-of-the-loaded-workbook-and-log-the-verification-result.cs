// Title: Verify an encrypted Excel workbook's password with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an .xlsx file using Aspose.Cells LoadOptions with a given password and returns a boolean indicating whether the password is correct. | Create a method that attempts to open a protected workbook, catches CellsException, and logs a success or failure message to the console. | Write a reusable function that accepts a file path and password, validates the password using Aspose.Cells, and outputs the verification result.
// Common Searches: Aspose.Cells C# check if Excel file password is correct | How to validate password for encrypted .xlsx using LoadOptions in Aspose.Cells | C# detect wrong password when opening protected workbook with Aspose.Cells | Example of catching CellsException for invalid Excel password Aspose.Cells | Verify workbook encryption password programmatically with Aspose.Cells .NET
// Tags: load encrypted xlsx with LoadOptions.Password Aspose.Cells | password validation for Excel workbook Aspose.Cells | catch CellsException invalid password | console logging password verification result | encrypted workbook handling Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

// The example loads an encrypted "encrypted.xlsx" file using Aspose.Cells LoadOptions.Password. If the Workbook loads without throwing a CellsException, the supplied password is reported as valid; otherwise, the code logs that the password is incorrect.
class Program
{
    static void Main()
    {
        // Path to the encrypted workbook
        string filePath = "encrypted.xlsx";

        // Verify that the file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        // The password to verify
        string passwordToCheck = "myPassword";

        bool isPasswordValid = false;

        try
        {
            // Attempt to load the workbook using the supplied password
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = passwordToCheck
            };

            // If loading succeeds, the password is correct
            Workbook workbook = new Workbook(filePath, loadOptions);
            isPasswordValid = true;
        }
        catch (CellsException)
        {
            // Loading failed – most likely due to an incorrect password
            isPasswordValid = false;
        }
        catch (Exception ex)
        {
            // Handle any other unexpected errors
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            return;
        }

        // Log the verification result
        if (isPasswordValid)
        {
            Console.WriteLine("Password verification succeeded: the password is correct.");
        }
        else
        {
            Console.WriteLine("Password verification failed: the password is incorrect.");
        }
    }
}
