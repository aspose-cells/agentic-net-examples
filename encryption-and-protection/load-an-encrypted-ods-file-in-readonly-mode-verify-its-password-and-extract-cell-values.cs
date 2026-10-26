// Title: Read a password‑protected ODS workbook in read‑only mode and extract its cell values using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells load encrypted ODS workbook with password | How to verify ODS file password using LoadOptions in Aspose.Cells .NET | Read cell values from a password‑protected ODS spreadsheet in C# | Aspose.Cells catch CellsException for incorrect ODS password | Iterate through worksheets and cells of a protected ODS file using Aspose.Cells
// Tags: encrypted ODS loading Aspose.Cells | ODS password verification LoadOptions | extract cell values from protected ODS | CellsException handling for ODS password | read‑only ODS workbook Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example checks for the encrypted ODS file, sets LoadOptions with the supplied password, and loads the workbook in read‑only mode using Aspose.Cells. It catches a CellsException to detect an incorrect password, then iterates through each worksheet and prints non‑null cell values.
    class Program
    {
        static void Main()
        {
            // Path to the encrypted ODS file
            string filePath = "encrypted.ods";

            // Password to verify
            string password = "myPassword";

            // Verify that the file exists before attempting to load it
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            // Configure load options for ODS format with the password
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Ods)
            {
                Password = password
            };

            Workbook workbook = null;

            try
            {
                // Load the encrypted ODS file
                workbook = new Workbook(filePath, loadOptions);
            }
            catch (CellsException ex)
            {
                // Aspose.Cells throws a CellsException when the password is incorrect
                if (ex.Message != null && ex.Message.IndexOf("Incorrect password", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    Console.WriteLine("Password verification failed.");
                }
                else
                {
                    Console.WriteLine($"Error loading file: {ex.Message}");
                }
                return;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
                return;
            }

            // Password is correct – extract and display cell values
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Console.WriteLine($"Sheet: {sheet.Name}");

                Cells cells = sheet.Cells;
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = cells[row, col];
                        if (cell.Value != null)
                        {
                            Console.WriteLine($"Cell {cell.Name}: {cell.Value}");
                        }
                    }
                }
            }
        }
    }
}
