// Title: How to Apply and Remove Password Write Protection on an Excel Workbook Using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that builds a new Excel workbook, inserts sample cells, applies write protection to the whole workbook (structure and windows) with a password via Aspose.Cells, and writes the file as a protected .xlsx. | Write C# code that opens a password‑protected .xlsx file, lifts the workbook protection using the same password through Aspose.Cells, and saves the resulting unprotected workbook.
// Common Searches: Aspose.Cells C# protect entire workbook with password and save as .xlsx | C# code to unprotect a password‑protected Excel file using Aspose.Cells Workbook.Unprotect | How to set workbook structure and windows protection in Aspose.Cells .NET | Remove write protection from an Excel workbook programmatically with the original password in C# | Example of applying and then removing password protection on an Excel workbook using Aspose.Cells
// Tags: Aspose.Cells workbook.Protect password protection | Aspose.Cells workbook.Unprotect usage | C# protect Excel workbook structure windows | C# remove Excel workbook write protection Aspose.Cells | save protected Excel .xlsx with Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a new workbook, adds data, protects the entire workbook (structure and windows) with a password using Workbook.Protect, saves it, then reloads the protected file, removes the protection with Workbook.Unprotect using the same password, and saves the unprotected workbook.
class Program
{
    static void Main()
    {
        try
        {
            // ---------- Create a new workbook ----------
            var workbook = new Workbook();

            // Add sample data
            var sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample");
            sheet.Cells["B1"].PutValue(42);

            // ---------- Apply write protection with a password ----------
            // Protect the entire workbook (structure and windows) with a password
            workbook.Protect(ProtectionType.All, "MyPassword");

            // Save the protected workbook
            const string protectedPath = "protected.xlsx";
            workbook.Save(protectedPath);
            Console.WriteLine($"Protected workbook saved to '{protectedPath}'.");

            // ---------- Load the protected workbook ----------
            if (!File.Exists(protectedPath))
                throw new FileNotFoundException($"The file '{protectedPath}' was not found.");

            var protectedWb = new Workbook(protectedPath);

            // ---------- Remove write protection using the same password ----------
            // Unprotect the workbook
            protectedWb.Unprotect("MyPassword");

            // Save the unprotected workbook
            const string unprotectedPath = "unprotected.xlsx";
            protectedWb.Save(unprotectedPath);
            Console.WriteLine($"Unprotected workbook saved to '{unprotectedPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
