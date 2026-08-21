// Title: Unprotect, edit, and re‑protect a worksheet with the same password using Aspose.Cells (.NET)
// Description: Demonstrates how to apply full protection to a worksheet, temporarily remove it, update cells, and then re‑apply the original password and protection options with Aspose.Cells for .NET.
// Keywords: Aspose.Cells worksheet unprotect | Aspose.Cells protect sheet C# | modify protected worksheet Aspose.Cells | re‑apply worksheet protection | worksheet protection options C# | Aspose.Cells Protect method | Aspose.Cells Unprotect method | C# Excel sheet security
// Common Searches: Aspose.Cells how to unprotect a worksheet in C# | C# change cell values in a protected sheet using Aspose.Cells | re‑protect worksheet with same password Aspose.Cells | preserve protection settings after editing worksheet Aspose.Cells | Aspose.Cells protect sheet with specific options
// Developer Intent: Temporarily lift worksheet protection, modify cell data, and restore the exact protection configuration using the original password.
// Use Cases: Bulk update data in a protected template while keeping the original security settings. | Automate editing of a locked worksheet and re‑apply the same protection before distribution. | Maintain consistent AllowEditingContent and AllowEditingObject flags after programmatic changes.
// AI Prompts: Generate C# code with Aspose.Cells that unprotects a worksheet, updates a range of cells, and re‑protects it using the same password and protection flags. | Explain how to capture the current Protection object settings before calling Unprotect and reuse them when calling Protect again. | Provide robust error‑handling patterns for worksheet protection and unprotection operations in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to apply full protection to a worksheet, temporarily remove it, update cells, and then re‑apply the original password and protection options with Aspose.Cells for .NET.
    public class UnprotectModifyReprotectDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // -----------------------------------------------------------------
                // Initial protection setup (store original options)
                // -----------------------------------------------------------------
                string password = "mySecretPassword";

                // Set desired protection options
                worksheet.Protection.AllowEditingContent = false;
                worksheet.Protection.AllowEditingObject = false;

                // Apply protection with the password (oldPassword is not required for new protection)
                worksheet.Protect(ProtectionType.All, password, null);

                // -----------------------------------------------------------------
                // Unprotect the worksheet using the stored password
                // -----------------------------------------------------------------
                worksheet.Unprotect(password);

                // -----------------------------------------------------------------
                // Modify cell values while the worksheet is unprotected
                // -----------------------------------------------------------------
                worksheet.Cells["A1"].PutValue("Updated Value");
                worksheet.Cells["B2"].PutValue(12345);

                // -----------------------------------------------------------------
                // Re‑protect the worksheet using the original password and options
                // -----------------------------------------------------------------
                worksheet.Protection.AllowEditingContent = false;
                worksheet.Protection.AllowEditingObject = false;

                // Apply protection again with the same password
                worksheet.Protect(ProtectionType.All, password, null);

                // -----------------------------------------------------------------
                // Save the workbook
                // -----------------------------------------------------------------
                workbook.Save("UnprotectedModifiedReprotected.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            UnprotectModifyReprotectDemo.Run();
        }
    }
}
