// Title: Copy a Range to a New Workbook and Apply Password Protection with Aspose.Cells for .NET (C#)
// Description: Loads a source workbook, copies the A1:B2 range to C3:D4 in a fresh workbook, sets a file‑level password, and saves the result as a protected Excel file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells copy range C# | Excel password protection .NET | copy cells to new workbook | file‑level encryption Aspose.Cells | C# Aspose.Cells example
// Common Searches: Aspose.Cells copy range between workbooks | set password on Excel file using Aspose.Cells | C# copy cells and protect workbook | how to encrypt an Excel file with Aspose.Cells | copy range and save as protected workbook
// Developer Intent: Transfer a specific cell block to a new workbook and secure the file with a password programmatically.
// Use Cases: Create a client‑ready report by extracting a table from a master workbook, placing it in a new file, and locking the file before distribution. | Automate generation of password‑protected templates where only selected data from a source sheet is copied to the template. | Build a batch process that copies data slices from multiple source files into individual protected workbooks for compliance purposes.
// AI Prompts: Write C# code with Aspose.Cells that copies a defined range from source.xlsx to destination.xlsx and adds a file password. | Show an Aspose.Cells .NET example that copies several non‑contiguous ranges into a new workbook and applies a password to the whole file. | Explain step‑by‑step how to handle missing source files when copying a range and protecting the destination workbook with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// Loads a source workbook, copies the A1:B2 range to C3:D4 in a fresh workbook, sets a file‑level password, and saves the result as a protected Excel file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the source workbook
            string sourcePath = "source.xlsx";

            // Ensure the source file exists; create a simple one if missing
            if (!File.Exists(sourcePath))
            {
                var tempWb = new Workbook();
                var tempWs = tempWb.Worksheets[0];
                tempWs.Cells["A1"].PutValue("Item1");
                tempWs.Cells["B1"].PutValue("Item2");
                tempWs.Cells["A2"].PutValue("Item3");
                tempWs.Cells["B2"].PutValue("Item4");
                tempWb.Save(sourcePath);
            }

            // Load the source workbook
            Workbook sourceWorkbook = new Workbook(sourcePath);

            // Create a new (empty) destination workbook
            Workbook destinationWorkbook = new Workbook();

            // Define the source range to copy (A1:B2 on the first worksheet)
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
            Aspose.Cells.Range sourceRange = sourceSheet.Cells.CreateRange("A1:B2");

            // Ensure the destination workbook has a worksheet to receive the data
            Worksheet destSheet = destinationWorkbook.Worksheets[0];

            // Define the destination range (C3:D4) where data will be pasted
            Aspose.Cells.Range destRange = destSheet.Cells.CreateRange("C3:D4");

            // Copy the source range into the destination range
            sourceRange.Copy(destRange);

            // Apply file‑level password protection to the entire workbook
            destinationWorkbook.Settings.Password = "MySecurePassword";

            // Save the protected workbook
            string destPath = "destination_protected.xlsx";
            destinationWorkbook.Save(destPath);

            Console.WriteLine($"Workbook saved successfully to '{destPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
