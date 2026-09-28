// Title: Copy a defined cell range to a new workbook and set a workbook‑level password with Aspose.Cells for .NET (C#)
// AI Prompts: Use Aspose.Cells in C# to copy the range A1:C10 from source.xlsx into a new workbook, assign the password 'MySecurePassword' to the new file, and save it as ProtectedCopy.xlsx. | Create a fresh Excel workbook, transfer a specific cell block from an existing workbook, then apply workbook‑wide password protection before saving, using the Aspose.Cells API. | Programmatically duplicate a range from one worksheet to another workbook and protect the resulting file with a password using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells copy range from one workbook to another and protect the new file with a password in C# | How to set workbook password after copying cells using Aspose.Cells .NET | C# example for copying A1:C10 to a new Excel file and applying password protection with Aspose.Cells | Save a copied range as a password‑protected Excel workbook using Aspose.Cells API | Aspose.Cells protect entire workbook after creating it from a cell range
// Tags: copy range to new workbook Aspose.Cells | workbook password protection .NET | Aspose.Cells copy cells between workbooks | C# create password‑protected Excel file | Aspose.Cells set Workbook.Settings.Password | export selected range as protected XLSX | Aspose.Cells range copy and file encryption

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// Loads source.xlsx, copies cells A1:C10 into a new workbook, assigns a password via Workbook.Settings.Password, and saves the result as ProtectedCopy.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            const string sourcePath = "source.xlsx";
            const string outputPath = "ProtectedCopy.xlsx";
            const string password = "MySecurePassword";

            // Verify that the source file exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the source workbook
            Workbook sourceWorkbook = new Workbook(sourcePath);

            // Create a new workbook and remove the default worksheet
            Workbook newWorkbook = new Workbook();
            newWorkbook.Worksheets.Clear();

            // Add a new worksheet to the new workbook
            Worksheet newSheet = newWorkbook.Worksheets.Add("Sheet1");

            // Define the range to copy from the source workbook
            const string sourceRange = "A1:C10";

            // Get source and destination ranges
            AsposeRange srcRange = sourceWorkbook.Worksheets[0].Cells.CreateRange(sourceRange);
            AsposeRange destRange = newSheet.Cells.CreateRange("A1");

            // Copy the defined range to the new worksheet starting at A1
            destRange.Copy(srcRange);

            // Apply password protection to the entire file
            newWorkbook.Settings.Password = password;

            // Save the new workbook with the applied protection
            newWorkbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
