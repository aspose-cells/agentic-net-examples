using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add sample data
        workbook.Worksheets[0].Cells["A1"].PutValue("Secure Data");

        // Set a modification (write) password
        workbook.Settings.WriteProtection.Password = "modify123";

        // Optional: set author and recommend read‑only flag
        workbook.Settings.WriteProtection.Author = "Admin";
        workbook.Settings.WriteProtection.RecommendReadOnly = true;

        // Save the workbook with write protection
        workbook.Save("ProtectedWorkbook.xlsx");

        // Load the saved workbook to verify protection
        Workbook loaded = new Workbook("ProtectedWorkbook.xlsx");
        bool isWriteProtected = loaded.Settings.WriteProtection.IsWriteProtected;
        Console.WriteLine("Is write protected: " + isWriteProtected);
    }
}