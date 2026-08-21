// Title: Copy a Cell Range to a New Workbook and Apply Password Protection with Aspose.Cells for .NET
// Description: Loads a source workbook, copies the range A1:C3 to a new workbook at E5:G7, encrypts the destination file with a password, and saves it as a protected Excel document using Aspose.Cells in C#.
// Keywords: Aspose.Cells copy range C# | new workbook password protection | encrypt Excel file Aspose.Cells | copy cells to another workbook .NET | Aspose.Cells workbook encryption example
// Common Searches: Aspose.Cells copy range to new workbook C# | how to password protect an Excel file with Aspose.Cells | copy cells A1:C3 to E5:G7 and encrypt workbook | C# Aspose.Cells example for range copy and protection | set workbook password after copying data Aspose
// Developer Intent: Transfer a specific cell block from an existing workbook into a fresh workbook and secure the resulting file with a password.
// Use Cases: Create a client‑ready report by extracting a table from a master file, placing it in a designated area of a new workbook, and locking the file before sending. | Automate compliance‑driven workflows that require only selected data to be shared in a password‑protected spreadsheet. | Generate template‑based dashboards where source data is copied into a fresh workbook and the output is encrypted to prevent unauthorized edits.
// AI Prompts: Generate C# code using Aspose.Cells to copy range A1:C3 from source.xlsx to E5:G7 in a new workbook and save it with a password. | Show an Aspose.Cells snippet that copies multiple ranges into a new workbook and applies workbook encryption with a custom password. | Explain step‑by‑step how to copy a cell range to another workbook and protect the file with a password using Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// Loads a source workbook, copies the range A1:C3 to a new workbook at E5:G7, encrypts the destination file with a password, and saves it as a protected Excel document using Aspose.Cells in C#.
class Program
{
    static void Main()
    {
        try
        {
            const string sourcePath = "source.xlsx";
            const string outputPath = "output_protected.xlsx";

            // Verify source file exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file \"{sourcePath}\" not found.");
                return;
            }

            // Load the source workbook from a file
            Workbook sourceWorkbook = new Workbook(sourcePath);

            // Create a new (empty) destination workbook
            Workbook destinationWorkbook = new Workbook();

            // Define the source range to copy (e.g., A1:C3 on the first worksheet)
            AsposeRange sourceRange = sourceWorkbook.Worksheets[0].Cells.CreateRange("A1:C3");

            // Define the destination range on the first worksheet of the new workbook
            // (e.g., start copying at cell E5, which will have the same size as the source range)
            Worksheet destSheet = destinationWorkbook.Worksheets[0];
            AsposeRange destinationRange = destSheet.Cells.CreateRange("E5:G7");

            // Copy the source range into the destination range (includes data, formatting, etc.)
            destinationRange.Copy(sourceRange);

            // Apply password protection to the entire workbook file (encryption)
            destinationWorkbook.Settings.Password = "SecretPassword123";

            // Save the protected workbook to a new file
            destinationWorkbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
