// Title: Copy a cell range from one workbook to a new workbook and set read‑only password protection on the sheet using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# copy A1:C10 range to a new workbook | How to protect an Aspose.Cells worksheet with a read‑only password in C# | Copy cells with formatting to another workbook using Aspose.Cells .NET | Set sheet protection type All with password using Aspose.Cells C# example | Create a new workbook from a selected range and apply password protection with Aspose.Cells
// Tags: range copy with style preservation Aspose.Cells | worksheet password protection read‑only Aspose.Cells | manual cell‑by‑cell copy Aspose.Cells C# | Protect method with ProtectionType.All example | export selected range to separate workbook C#

using System;
using System.IO;
using Aspose.Cells;

// The code loads source.xlsx, manually copies the A1:C10 range (including values and styles) from the first worksheet to a newly created workbook's "CopiedData" sheet, applies read‑only protection to that sheet with a password using Protect(ProtectionType.All,…), and saves the result as dest.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the source workbook
            string sourcePath = "source.xlsx";

            // Verify that the source file exists
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the source workbook
            Workbook sourceWorkbook = new Workbook(sourcePath);
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

            // Define the range to copy (A1:C10)
            CellArea sourceArea = new CellArea
            {
                StartRow = 0,    // Row 1
                StartColumn = 0, // Column A
                EndRow = 9,      // Row 10
                EndColumn = 2    // Column C
            };

            // Create a new workbook and add a fresh worksheet
            Workbook destWorkbook = new Workbook();
            destWorkbook.Worksheets.Clear();
            Worksheet destSheet = destWorkbook.Worksheets.Add("CopiedData");

            // Manually copy the defined range from source to destination sheet
            try
            {
                for (int row = sourceArea.StartRow; row <= sourceArea.EndRow; row++)
                {
                    for (int col = sourceArea.StartColumn; col <= sourceArea.EndColumn; col++)
                    {
                        Cell srcCell = sourceSheet.Cells[row, col];
                        Cell destCell = destSheet.Cells[row - sourceArea.StartRow, col - sourceArea.StartColumn];

                        // Copy value
                        destCell.PutValue(srcCell.Value);

                        // Copy style
                        Style srcStyle = srcCell.GetStyle();
                        destCell.SetStyle(srcStyle);
                    }
                }
            }
            catch (Exception copyEx)
            {
                Console.WriteLine($"Error during range copy: {copyEx.Message}");
                return;
            }

            // Protect the destination sheet (read‑only) with a password
            string password = "myPassword";
            destSheet.Protect(ProtectionType.All, password, null);

            // Save the new workbook
            string destPath = "dest.xlsx";
            destWorkbook.Save(destPath);
            Console.WriteLine($"Workbook saved successfully to {destPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
