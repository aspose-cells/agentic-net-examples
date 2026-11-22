// Title: C# example: attempt to edit a password‑protected worksheet with Aspose.Cells and capture the resulting CellsException
// Common Searches: Aspose.Cells C# edit cell on password protected worksheet without providing password | how to catch CellsException when modifying a locked sheet using Aspose.Cells | record failure of unauthorized edit on a protected Excel worksheet with Aspose.Cells .NET | Aspose.Cells attempt to write to a protected sheet and detect error in C# | C# Aspose.Cells protect worksheet then test edit failure
// Tags: Aspose.Cells worksheet protection C# | CellsException handling Aspose.Cells | modify protected Excel sheet without password | Aspose.Cells ProtectionType.All usage | unauthorized edit failure recording Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The sample creates a new Workbook, protects the first worksheet with a password using sheet.Protect(ProtectionType.All, "pwd123", null), then tries to write to cell A1 without supplying the password. The operation throws a CellsException, which is caught and logged, demonstrating how to detect and record an unauthorized edit attempt on a protected worksheet before saving the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Protect the worksheet with a password (oldPassword not required for new protection)
            sheet.Protect(ProtectionType.All, "pwd123", null);

            // Attempt to edit a cell without providing the password
            try
            {
                // This operation should fail because the sheet is protected
                sheet.Cells["A1"].PutValue("Attempted Edit");
                Console.WriteLine("Edit succeeded unexpectedly.");
            }
            catch (CellsException ex)
            {
                Console.WriteLine("Failed to edit protected worksheet: " + ex.Message);
            }

            // Save the workbook
            string outputPath = "ProtectedSheet.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An unexpected error occurred: " + ex.Message);
        }
    }
}
