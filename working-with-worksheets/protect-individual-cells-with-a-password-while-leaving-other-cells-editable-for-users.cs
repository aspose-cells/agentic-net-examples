// Title: Protect Specific Cells with a Password in Aspose.Cells for .NET (C#)
// Description: Demonstrates how to lock individual cells with a password while leaving other cells editable using Aspose.Cells for .NET. The example creates a workbook, writes sample data, unlocks a range (B1:C1) by applying a style with IsLocked = false, adds a password‑protected range for cell A1 via AllowEditRanges, optionally secures the whole sheet, and saves the file as ProtectedIndividualCells.xlsx.
// Keywords: Aspose.Cells protect cell password | C# lock single cell Excel | AllowEditRanges Aspose.Cells | unlock cell range Aspose.Cells | worksheet protection .NET | password protected Excel cell C# | Aspose.Cells cell level security
// Common Searches: how to lock a single cell with a password using Aspose.Cells | unlock a range of cells while protecting the rest of a worksheet in C# | set password for specific cells in an Excel file with Aspose.Cells | cell‑level protection Aspose.Cells .NET example | protect individual cells Aspose.Cells C# tutorial
// Developer Intent: The developer needs to apply password protection to selected cells while keeping the remaining cells editable.
// Use Cases: Create a template where only input cells are editable and calculation cells are password‑locked. | Distribute a financial model that safeguards key formulas but allows users to modify assumptions. | Generate a report that hides confidential values behind cell‑level passwords while exposing summary fields for editing.
// AI Prompts: Show how to protect multiple non‑contiguous cells with different passwords using Aspose.Cells for .NET. | Explain how to check if a cell is locked and retrieve its password programmatically with Aspose.Cells. | Provide code to change the password of an existing protected range without affecting other protections.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to lock individual cells with a password while leaving other cells editable using Aspose.Cells for .NET. The example creates a workbook, writes sample data, unlocks a range (B1:C1) by applying a style with IsLocked = false, adds a password‑protected range for cell A1 via AllowEditRanges, optionally secures the whole sheet, and saves the file as ProtectedIndividualCells.xlsx.
    public class ProtectIndividualCellsDemo
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Fill some sample data
            cells["A1"].PutValue("Password Protected Cell");
            cells["B1"].PutValue("Editable Cell 1");
            cells["C1"].PutValue("Editable Cell 2");
            cells["A2"].PutValue("Another Locked Cell");

            // Unlock the range B1:C1 so users can edit without a password
            Style unlockedStyle = workbook.CreateStyle();
            unlockedStyle.IsLocked = false;
            StyleFlag flag = new StyleFlag();
            flag.Locked = true; // apply the lock property from the style
            cells.CreateRange("B1:C1").ApplyStyle(unlockedStyle, flag);

            // Add a protected range for cell A1 that requires a password to edit
            int rangeIndex = worksheet.AllowEditRanges.Add("PasswordProtectedA1", 0, 0, 0, 0);
            ProtectedRange protectedRange = worksheet.AllowEditRanges[rangeIndex];
            protectedRange.Password = "cellpwd";

            // Protect the worksheet with a sheet password (optional)
            worksheet.Protect(ProtectionType.All, "sheetpwd", null);

            // Save the workbook
            string outputPath = "ProtectedIndividualCells.xlsx";
            workbook.Save(outputPath);

            // Output verification information
            Console.WriteLine("Worksheet protected with password: " + worksheet.Protection.IsProtectedWithPassword);
            Console.WriteLine("Cell A1 requires password: " + protectedRange.IsProtectedWithPassword);
            Console.WriteLine("Editable range B1:C1 is unlocked.");
            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
    }
}
