// Title: C# – Unprotect an Excel worksheet by testing common passwords with Aspose.Cells
// Description: The routine loads a workbook, selects a worksheet, and iterates a predefined list of common passwords. For each entry it calls worksheet.Protection.VerifyPassword to test validity without raising an exception; when a match is found it invokes worksheet.Unprotect and finally saves the (potentially) unprotected file.
// Keywords: Aspose.Cells | C# | Excel worksheet unprotect | common password list | VerifyPassword | worksheet protection API | programmatic Excel de‑protection | load workbook | save workbook | brute‑force password Excel
// Common Searches: how to unprotect an Excel worksheet programmatically using Aspose.Cells | C# code to try common passwords on a protected worksheet | Aspose.Cells VerifyPassword example | remove worksheet protection without knowing the password | save workbook after attempting to unprotect sheet in .NET
// Developer Intent: Automatically attempt to remove worksheet protection by testing a set of typical passwords and persist the outcome.
// Use Cases: Regain access to a sheet when the original password is forgotten but likely weak. | Batch‑process a collection of workbooks to strip trivial worksheet protection before data extraction. | Integrate into a migration tool that cleans protected sheets prior to format conversion.
// AI Prompts: Generate C# code using Aspose.Cells that tries a custom password list to unprotect a worksheet and logs the successful password. | Explain the behavior of worksheet.Protection.VerifyPassword in Aspose.Cells and how to handle verification errors during password trials. | Show how to extend the example to loop through all worksheets in a workbook and apply the common‑password test to each.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The routine loads a workbook, selects a worksheet, and iterates a predefined list of common passwords. For each entry it calls worksheet.Protection.VerifyPassword to test validity without raising an exception; when a match is found it invokes worksheet.Unprotect and finally saves the (potentially) unprotected file.
    public class WorksheetUnprotectHelper
    {
        // List of common passwords to try
        private static readonly string[] CommonPasswords = new string[]
        {
            "123456", "password", "admin", "test", "12345",
            "1234", "123", "password1", "12345678", "qwerty"
        };

        /// <param name="inputFilePath">Path to the protected workbook.</param>
        /// <param name="outputFilePath">Path where the (potentially) unprotected workbook will be saved.</param>
        public static void UnprotectWorksheetWithCommonPasswords(string inputFilePath, string outputFilePath)
        {
            if (!File.Exists(inputFilePath))
            {
                Console.WriteLine($"Input file not found: {inputFilePath}");
                return;
            }

            try
            {
                // Load the workbook from the specified file
                Workbook workbook = new Workbook(inputFilePath);

                // Access the first worksheet (adjust index as needed)
                Worksheet worksheet = workbook.Worksheets[0];

                bool unprotected = false;

                // Iterate through the list of common passwords
                foreach (string pwd in CommonPasswords)
                {
                    try
                    {
                        // Verify the password without throwing an exception
                        if (worksheet.Protection.VerifyPassword(pwd))
                        {
                            // Correct password found – unprotect the worksheet
                            worksheet.Unprotect(pwd);
                            unprotected = true;
                            Console.WriteLine($"Worksheet unprotected using password: \"{pwd}\"");
                            break;
                        }
                    }
                    catch (Exception verifyEx)
                    {
                        // Log verification errors but continue trying other passwords
                        Console.WriteLine($"Error verifying password \"{pwd}\": {verifyEx.Message}");
                    }
                }

                if (!unprotected)
                {
                    Console.WriteLine("Failed to unprotect the worksheet with the provided common passwords or worksheet is not protected.");
                }

                // Save the workbook (whether modified or not) to the output path
                try
                {
                    workbook.Save(outputFilePath);
                    Console.WriteLine($"Workbook saved to: {outputFilePath}");
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine($"Error saving workbook: {saveEx.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while processing the workbook: {ex.Message}");
            }
        }
    }

    public class Program
    {
        // Entry point of the console application
        public static void Main(string[] args)
        {
            // Example file paths – adjust as needed or pass via command‑line arguments
            string inputPath = "protected.xlsx";
            string outputPath = "unprotected.xlsx";

            if (args.Length >= 2)
            {
                inputPath = args[0];
                outputPath = args[1];
            }

            WorksheetUnprotectHelper.UnprotectWorksheetWithCommonPasswords(inputPath, outputPath);
        }
    }
}
