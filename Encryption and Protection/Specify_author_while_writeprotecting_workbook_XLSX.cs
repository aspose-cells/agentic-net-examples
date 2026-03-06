using System;
using Aspose.Cells;

namespace AsposeCellsWriteProtectionDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access write‑protection settings via Workbook.Settings.WriteProtection
            WriteProtection wp = workbook.Settings.WriteProtection;

            // Set the author who applied the write protection
            wp.Author = "John Doe";

            // Set a password required to modify the file
            wp.Password = "securePassword123";

            // Optionally recommend opening the file as read‑only
            wp.RecommendReadOnly = true;

            // Save the workbook (lifecycle: save)
            string outputPath = "WriteProtectedWorkbook.xlsx";
            workbook.Save(outputPath);

            // Load the saved workbook to verify the settings (lifecycle: load)
            Workbook loaded = new Workbook(outputPath);
            Console.WriteLine("Author: " + loaded.Settings.WriteProtection.Author);
            Console.WriteLine("Is Write Protected: " + loaded.Settings.WriteProtection.IsWriteProtected);
            Console.WriteLine("Recommend Read‑Only: " + loaded.Settings.WriteProtection.RecommendReadOnly);
        }
    }
}