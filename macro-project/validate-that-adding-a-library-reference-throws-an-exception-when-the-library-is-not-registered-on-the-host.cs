// Title: How to verify that Aspose.Cells throws an exception when loading a non‑existent Excel workbook in C#
// Common Searches: Aspose.Cells C# catch exception for missing workbook file | load Excel file with Aspose.Cells throws file not found error | unit test Aspose.Cells workbook constructor with invalid path | C# verify exception is raised for non‑existent .xlsx using Aspose.Cells | Aspose.Cells handling FileNotFoundException when opening workbook
// Tags: Aspose.Cells workbook load missing file exception | C# Aspose.Cells file not found handling | validate workbook constructor error Aspose.Cells | exception handling for absent Excel workbook Aspose.Cells | test missing workbook path Aspose.Cells C#

using System;
using Aspose.Cells;
using System.IO;

// The sample attempts to create a Workbook with a path that does not exist, expects an exception, and prints a pass/fail message based on whether the exception is caught, demonstrating proper error handling for missing Excel files in Aspose.Cells.
class Program
{
    static void Main()
    {
        try
        {
            // Attempt to load a workbook from a non‑existent file.
            // This should throw an exception because the file does not exist.
            string fakePath = "NonExistentLibrary.xlsx";

            // Ensure we don't attempt to load if the file somehow exists.
            if (File.Exists(fakePath))
            {
                // If the file unexpectedly exists, delete it to keep the test valid.
                File.Delete(fakePath);
            }

            // This constructor will raise an exception for a missing file.
            Workbook workbook = new Workbook(fakePath);

            // If no exception is thrown, the test has failed.
            Console.WriteLine("Test Failed: No exception was thrown for a missing workbook file.");
        }
        catch (Exception ex)
        {
            // Expected outcome: an exception is thrown.
            Console.WriteLine("Test Passed: Caught expected exception -> " + ex.Message);
        }
    }
}
