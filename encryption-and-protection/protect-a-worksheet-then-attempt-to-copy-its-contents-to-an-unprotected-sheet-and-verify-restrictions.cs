// Title: Aspose.Cells for .NET: Protect a Worksheet, Copy Its Data to an Unprotected Sheet, and Verify Protection Behavior
// Description: Demonstrates how to protect the first worksheet of a workbook with a password, persist the protection after saving, export its values, import them into a newly added unprotected sheet, test edit restrictions on the protected sheet, unprotect it, modify a cell, and save the final file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells worksheet protection | C# protect Excel sheet password | copy data from protected worksheet | ExportArray Aspose.Cells | ImportArray Aspose.Cells | IsProtected property | Unprotect worksheet C# | Excel sheet security .NET | persist worksheet protection | copy values between sheets Aspose | modify protected cell Aspose.Cells | Aspose.Cells lifecycle rule
// Common Searches: protect Excel worksheet with password using Aspose.Cells C# | copy cells from a protected sheet to another sheet Aspose.Cells | check if worksheet protection is retained after saving | unprotect worksheet and edit cells Aspose.Cells .NET | ExportArray and ImportArray example Aspose.Cells
// Developer Intent: The developer needs to lock a worksheet with a password, duplicate its content to a separate unprotected sheet, confirm that protected cells cannot be edited until the sheet is unlocked, and then save the workbook.
// Use Cases: Ensuring worksheet protection survives file save/load cycles. | Migrating data from a locked sheet to a new sheet without transferring protection settings. | Programmatically validating that protected cells reject modifications until unprotected. | Demonstrating the ExportArray/ImportArray workflow for protected worksheets.
// AI Prompts: Generate C# code with Aspose.Cells that protects a worksheet, copies its values to another sheet, and verifies protection status before and after unprotecting. | Explain why Aspose.Cells allows PutValue on a protected cell and how to enforce true read‑only behavior. | Show how to use IsProtected and Unprotect methods to test worksheet security in a .NET application.

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to protect the first worksheet of a workbook with a password, persist the protection after saving, export its values, import them into a newly added unprotected sheet, test edit restrictions on the protected sheet, unprotect it, modify a cell, and save the final file using Aspose.Cells for .NET.
class WorksheetProtectionDemo
{
    static void Main()
    {
        try
        {
            // -------------------------------------------------
            // 1. Create a new workbook and add some data
            // -------------------------------------------------
            Workbook wb = new Workbook();
            Worksheet src = wb.Worksheets[0];
            src.Cells["A1"].PutValue("Hello");
            src.Cells["B2"].PutValue(123);

            // -------------------------------------------------
            // 2. Protect the source worksheet with a password
            // -------------------------------------------------
            src.Protect(ProtectionType.All, "pwd123", null);
            Console.WriteLine("Source sheet protected (in memory): " + src.IsProtected);

            // -------------------------------------------------
            // 3. Save the workbook (lifecycle rule)
            // -------------------------------------------------
            string protectedPath = "protected.xlsx";
            wb.Save(protectedPath);

            // -------------------------------------------------
            // 4. Load the workbook to verify protection persists
            // -------------------------------------------------
            if (!File.Exists(protectedPath))
            {
                Console.WriteLine($"File not found: {protectedPath}");
                return;
            }

            Workbook loadedWb = new Workbook(protectedPath);
            Worksheet loadedSrc = loadedWb.Worksheets[0];
            Console.WriteLine("Loaded source sheet protected: " + loadedSrc.IsProtected);

            // -------------------------------------------------
            // 5. Add a new (unprotected) worksheet as destination
            // -------------------------------------------------
            int destIndex = loadedWb.Worksheets.Add();
            Worksheet dest = loadedWb.Worksheets[destIndex];

            // -------------------------------------------------
            // 6. Copy the contents from the protected sheet to the unprotected sheet
            // -------------------------------------------------
            int totalRows = loadedSrc.Cells.MaxDisplayRange.RowCount;
            int totalCols = loadedSrc.Cells.MaxDisplayRange.ColumnCount;

            // Export values from source
            object[,] values = loadedSrc.Cells.ExportArray(0, 0, totalRows, totalCols);

            // Convert to string[,] because ImportArray overload expects string[,] in this SDK version
            string[,] stringValues = new string[totalRows, totalCols];
            for (int i = 0; i < totalRows; i++)
            {
                for (int j = 0; j < totalCols; j++)
                {
                    object val = values[i, j];
                    stringValues[i, j] = val?.ToString() ?? string.Empty;
                }
            }

            // Import into destination worksheet
            dest.Cells.ImportArray(stringValues, 0, 0);
            Console.WriteLine("Destination sheet protected: " + dest.IsProtected);

            // -------------------------------------------------
            // 7. Attempt to modify a cell in the protected sheet without unprotecting
            // -------------------------------------------------
            try
            {
                loadedSrc.Cells["A1"].PutValue("Modified without unprotect");
                Console.WriteLine("Modified protected sheet without unprotect (API permits it).");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error modifying protected sheet: " + ex.Message);
            }

            // -------------------------------------------------
            // 8. Unprotect the worksheet using the correct password
            // -------------------------------------------------
            loadedSrc.Unprotect("pwd123");
            Console.WriteLine("Source sheet protected after unprotect: " + loadedSrc.IsProtected);

            // -------------------------------------------------
            // 9. Modify the cell after unprotecting to show normal operation
            // -------------------------------------------------
            loadedSrc.Cells["A1"].PutValue("Modified after unprotect");
            Console.WriteLine("Cell A1 after modification: " + loadedSrc.Cells["A1"].StringValue);

            // -------------------------------------------------
            // 10. Save the final workbook (lifecycle rule)
            // -------------------------------------------------
            string finalPath = "final.xlsx";
            loadedWb.Save(finalPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Unexpected error: " + ex.Message);
        }
    }
}
