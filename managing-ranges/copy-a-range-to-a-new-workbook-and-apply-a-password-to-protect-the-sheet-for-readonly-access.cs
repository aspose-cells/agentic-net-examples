// Title: Copy a Range to a New Workbook and Apply a Read‑Only Password with Aspose.Cells (C#)
// Description: Loads a source workbook, defines a cell range (e.g., A1:C5), copies that range into an empty workbook, protects the destination worksheet with a read‑only password, and saves the result as a separate file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells copy range | C# copy range to new workbook | Aspose.Cells worksheet protection | read‑only password Excel .NET | Range.Copy Aspose.Cells | Protect worksheet Aspose.Cells | Excel automation C# | Aspose.Cells .NET example
// Common Searches: copy range from one Excel file to another Aspose.Cells C# | set read‑only password on worksheet Aspose.Cells | how to protect copied sheet with password using Aspose.Cells | Aspose.Cells example copy cells and lock sheet | C# Aspose.Cells protect worksheet read only
// Developer Intent: Copy a defined cell block from an existing workbook into a new workbook and lock the new sheet with a password that permits only read‑only access.
// Use Cases: Create a distribution‑ready report by extracting a data block from a master file and preventing edits. | Generate a snapshot of a summary range for archival, ensuring the copied content cannot be modified. | Automate the production of a protected workbook for external stakeholders who need view‑only access.
// AI Prompts: Write C# code with Aspose.Cells that copies range A1:C5 from source.xlsx to a new workbook and protects the sheet with password 'ReadOnlyPwd' for read‑only access. | Show an Aspose.Cells example that copies a named range, handles missing source files, and applies worksheet protection with a custom password. | Explain how to copy multiple ranges into separate worksheets and assign different passwords to each using Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// Loads a source workbook, defines a cell range (e.g., A1:C5), copies that range into an empty workbook, protects the destination worksheet with a read‑only password, and saves the result as a separate file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            const string sourcePath = "source.xlsx";
            const string destPath = "copied_protected.xlsx";

            // Verify source file exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the source workbook
            Workbook sourceWorkbook = new Workbook(sourcePath);
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

            // Define the range to copy (e.g., A1:C5)
            int startRow = 0;          // Row index (0‑based)
            int startColumn = 0;       // Column index (0‑based)
            int rowCount = 5;          // Number of rows in the range
            int columnCount = 3;       // Number of columns in the range

            // Create the source range object
            AsposeRange sourceRange = sourceSheet.Cells.CreateRange(startRow, startColumn, rowCount, columnCount);

            // Create a new (empty) workbook for the destination
            Workbook destinationWorkbook = new Workbook();
            Worksheet destinationSheet = destinationWorkbook.Worksheets[0];

            // Create a destination range of the same size
            AsposeRange destinationRange = destinationSheet.Cells.CreateRange(startRow, startColumn, rowCount, columnCount);

            // Copy the source range into the destination range
            sourceRange.Copy(destinationRange);

            // Protect the destination worksheet with a password (read‑only access)
            destinationSheet.Protect(ProtectionType.All, "ReadOnlyPwd", null);

            // Save the new workbook
            destinationWorkbook.Save(destPath);
            Console.WriteLine($"Workbook saved successfully to {destPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
