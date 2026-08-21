// Title: Copy a Cell Range to a New Workbook and Protect Its Structure with a Password – Aspose.Cells for .NET (C#)
// Description: C# example that loads or creates a source Excel file, defines range A1:C5, copies it into a fresh workbook, applies structure protection with a password, and saves the result as copied_protected.xlsx using Aspose.Cells.
// Keywords: Aspose.Cells copy range C# | protect workbook structure password | copy cells between Excel files .NET | Aspose.Cells workbook protection | C# Excel range export | Aspose.Cells create workbook from range
// Common Searches: Aspose.Cells copy range to another workbook C# | How to protect workbook structure with password using Aspose.Cells | Copy A1:C5 to new Excel file and lock structure | C# Aspose.Cells example for range copy and protection
// Developer Intent: Transfer a specific cell range from an existing workbook to a newly created workbook and then secure the new workbook’s sheet structure with a password.
// Use Cases: Generate a template by extracting a data block from a master workbook and preventing sheet reordering. | Create a read‑only report that contains only the calculation area while locking the workbook structure. | Automate exporting selected ranges to separate files for distribution, ensuring recipients cannot add, delete, or move sheets.
// AI Prompts: Write C# code with Aspose.Cells that copies range A1:C5 from source.xlsx to a new workbook and protects the workbook structure with password "myPassword". | Explain how to copy any range between two Excel workbooks using Aspose.Cells and then apply structure protection, including handling a missing source file. | Provide a step‑by‑step guide to copy multiple ranges into different worksheets of a new workbook and set distinct structure passwords for each workbook with Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;

// C# example that loads or creates a source Excel file, defines range A1:C5, copies it into a fresh workbook, applies structure protection with a password, and saves the result as copied_protected.xlsx using Aspose.Cells.
class Program
{
    static void Main()
    {
        try
        {
            const string sourcePath = "source.xlsx";
            const string destPath = "copied_protected.xlsx";

            // Ensure source file exists; create a simple one if missing
            if (!File.Exists(sourcePath))
            {
                var tempWb = new Workbook();
                var tempSheet = tempWb.Worksheets[0];
                // Fill A1:C5 with sample data
                for (int row = 0; row < 5; row++)
                {
                    for (int col = 0; col < 3; col++)
                    {
                        tempSheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                    }
                }
                tempWb.Save(sourcePath, SaveFormat.Xlsx);
            }

            // Load the source workbook containing the range to copy
            Workbook sourceWorkbook = new Workbook(sourcePath);

            // Access the first worksheet and define the source range (e.g., A1:C5)
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
            Aspose.Cells.Range sourceRange = sourceSheet.Cells.CreateRange("A1:C5");

            // Create a new (empty) workbook that will receive the copied range
            Workbook destinationWorkbook = new Workbook();

            // Ensure the destination workbook has a worksheet to receive data
            Worksheet destinationSheet = destinationWorkbook.Worksheets[0];

            // Create a destination range of the same size starting at A1
            Aspose.Cells.Range destinationRange = destinationSheet.Cells.CreateRange("A1:C5");

            // Copy the source range into the destination range
            destinationRange.Copy(sourceRange);

            // Protect the workbook structure with a password
            destinationWorkbook.Protect(ProtectionType.Structure, "myPassword");

            // Save the new workbook
            destinationWorkbook.Save(destPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved successfully to '{destPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
