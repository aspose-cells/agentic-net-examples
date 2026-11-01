// Title: Validate VBA project password length before calling Workbook.VbaProject.Protect with Aspose.Cells for .NET (C#)
// Common Searches: how to enforce minimum password length when protecting a VBA project using Aspose.Cells in C# | C# code to validate VBA project password before invoking Workbook.VbaProject.Protect | example of password length requirement for Aspose.Cells VBA project protection | throw exception if VBA project password is shorter than required with Aspose.Cells | check VBA password size before protecting Excel workbook via Aspose.Cells .NET
// Tags: Aspose.Cells VBA password validation | Workbook.VbaProject.Protect usage C# | minimum VBA password length Aspose.Cells | protect VBA project programmatically .NET | Excel workbook VBA protection password check

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing Excel file, defines a VBA project password, verifies that the password is not null and meets a minimum length of eight characters, throws an ArgumentException if the check fails, then protects the VBA project with the validated password using Workbook.VbaProject.Protect and saves the workbook, handling any errors that occur.
class VbaPasswordProtection
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Define the VBA project password
            string vbaPassword = "YourSecurePassword";

            // Minimum password length requirement
            const int MinPasswordLength = 8;

            // Validate password length before protecting the VBA project
            if (string.IsNullOrEmpty(vbaPassword) || vbaPassword.Length < MinPasswordLength)
            {
                throw new ArgumentException($"VBA project password must be at least {MinPasswordLength} characters long.");
            }

            // Protect the VBA project (false = not read‑only, password provided)
            workbook.VbaProject.Protect(false, vbaPassword);

            // Save the workbook with the protected VBA project
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
