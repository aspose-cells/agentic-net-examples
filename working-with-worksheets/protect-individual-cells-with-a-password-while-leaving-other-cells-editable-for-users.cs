// Title: C# – Protect Specific Cells with a Password in Aspose.Cells for .NET
// Description: Creates a workbook, defines an AllowEditRange (e.g., A1:B2), assigns a password, locks that range, unlocks other cells (e.g., C1:D5), protects the worksheet (no sheet‑level password), and saves the file as IndividualCellPasswordProtection.xlsx.
// Keywords: Aspose.Cells C# protect cells | cell password protection Aspose | AllowEditRanges .NET | lock/unlock cells Aspose.Cells | worksheet protection without sheet password | individual cell password Aspose
// Common Searches: How to set a password for a cell range using Aspose.Cells C# | Aspose.Cells protect some cells and leave others editable | C# code to lock cells with password in Excel via Aspose | AllowEditRanges example Aspose.Cells | Worksheet protection with cell‑level passwords .NET
// Developer Intent: Add a password‑protected range to an Excel worksheet while keeping the rest of the cells editable, using Aspose.Cells in C#.
// Use Cases: Financial templates where total cells are locked and require a password to modify, but input cells stay editable. | Shared spreadsheets that allow collaborators to fill data entry fields while safeguarding formula cells with a password. | Protecting confidential information in specific cells of a distributed workbook while permitting edits elsewhere.
// AI Prompts: Generate C# code to protect multiple non‑contiguous cell ranges with different passwords using Aspose.Cells. | Show how to lock all cells by default, then unlock a dynamic list of cells based on user input, and finally protect the worksheet. | Explain how to change or remove the password of an existing AllowEditRange in a saved workbook.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a workbook, defines an AllowEditRange (e.g., A1:B2), assigns a password, locks that range, unlocks other cells (e.g., C1:D5), protects the worksheet (no sheet‑level password), and saves the file as IndividualCellPasswordProtection.xlsx.
    public class ProtectIndividualCellsWithPassword
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // -------------------------------------------------
                // 1. Define the range that should be protected by a password (e.g., A1:B2)
                // -------------------------------------------------
                // Add the range to the AllowEditRanges collection
                // The range is initially allowed to edit; we will set a password so that editing requires it
                int protectedRangeIndex = worksheet.AllowEditRanges.Add("PasswordProtectedRange", 0, 0, 1, 1);
                ProtectedRange protectedRange = worksheet.AllowEditRanges[protectedRangeIndex];
                protectedRange.Password = "CellPass123"; // password required to edit this range

                // -------------------------------------------------
                // 2. Ensure the cells in the protected range are locked
                // -------------------------------------------------
                // Locking has no effect unless the worksheet is protected, but we set it explicitly
                for (int row = 0; row <= 1; row++)
                {
                    for (int col = 0; col <= 1; col++)
                    {
                        Cell cell = cells[row, col];
                        Style style = cell.GetStyle();
                        style.IsLocked = true;
                        cell.SetStyle(style);
                    }
                }

                // -------------------------------------------------
                // 3. Define the range that should remain editable (e.g., C1:D5)
                // -------------------------------------------------
                for (int row = 0; row <= 4; row++)
                {
                    for (int col = 2; col <= 3; col++)
                    {
                        Cell cell = cells[row, col];
                        Style style = cell.GetStyle();
                        style.IsLocked = false; // unlock these cells
                        cell.SetStyle(style);
                    }
                }

                // -------------------------------------------------
                // 4. Protect the worksheet (no sheet‑level password)
                // -------------------------------------------------
                // This enforces the lock settings and activates the password on the protected range
                worksheet.Protect(ProtectionType.All);

                // -------------------------------------------------
                // 5. Save the workbook
                // -------------------------------------------------
                workbook.Save("IndividualCellPasswordProtection.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                ProtectIndividualCellsWithPassword.Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
