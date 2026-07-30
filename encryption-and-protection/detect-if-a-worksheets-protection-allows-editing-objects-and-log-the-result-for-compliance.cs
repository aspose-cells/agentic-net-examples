// Title: Check Worksheet AllowEditingObject Protection with Aspose.Cells for .NET (C#)
// Description: C# example that loads an Excel workbook (creates a placeholder if missing), reads the worksheet's Protection.AllowEditingObject flag, logs the result for compliance auditing, and saves the file unchanged.
// Keywords: Aspose.Cells | C# worksheet protection | AllowEditingObject | read protection flag | Excel compliance | detect object editing | worksheet security | Aspose.Cells .NET | Excel protection API
// Common Searches: Aspose.Cells read AllowEditingObject | C# check worksheet protection object editing | How to get worksheet AllowEditingObject flag | Aspose.Cells compliance audit Excel | Detect if objects can be edited in Excel sheet using Aspose
// Developer Intent: Identify whether a worksheet permits editing of drawing objects by reading the AllowEditingObject property and log the outcome.
// Use Cases: Compliance check before processing user‑uploaded spreadsheets | Generate a security report of all sheets in a workbook | Fail a CI build if any sheet allows object editing | Automate audit of Excel templates for object‑editing restrictions
// AI Prompts: Provide a C# loop that scans every worksheet in a workbook and returns the names of sheets where AllowEditingObject is true using Aspose.Cells. | Show how to disable AllowEditingObject, protect the worksheet with a password, and save the workbook in C#. | Create a PowerShell script that uses Aspose.Cells to export the AllowEditingObject status of each worksheet to a CSV file.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // C# example that loads an Excel workbook (creates a placeholder if missing), reads the worksheet's Protection.AllowEditingObject flag, logs the result for compliance auditing, and saves the file unchanged.
    public class DetectEditingObjectsCompliance
    {
        // Entry point for the application
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        public static void Run()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            Workbook workbook;

            // Ensure the input file exists; if not, create a new workbook
            if (File.Exists(inputPath))
            {
                try
                {
                    workbook = new Workbook(inputPath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to load workbook '{inputPath}': {ex.Message}");
                    return;
                }
            }
            else
            {
                Console.WriteLine($"Input file '{inputPath}' not found. Creating a new workbook.");
                workbook = new Workbook(); // creates a default workbook with one worksheet
                workbook.Save(inputPath); // optionally save the placeholder file
            }

            // Access the first worksheet (adjust index as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Retrieve the protection settings for the worksheet
            Protection protection = worksheet.Protection;

            // Check whether editing of drawing objects is permitted
            bool allowEditingObject = protection.AllowEditingObject;

            // Log the compliance result
            Console.WriteLine($"Worksheet '{worksheet.Name}' AllowEditingObject: {allowEditingObject}");

            // Save the workbook (no modifications made)
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook '{outputPath}': {ex.Message}");
            }
        }
    }
}
