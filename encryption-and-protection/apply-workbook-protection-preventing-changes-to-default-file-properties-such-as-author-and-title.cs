// Title: Protect Excel workbook metadata (author, title) with write‑protection using Aspose.Cells for .NET
// Description: Shows how to configure Aspose.Cells Workbook.Settings.WriteProtection to set an author, assign a password, recommend read‑only mode, save the file, and verify the protection flags in C#.
// Keywords: Aspose.Cells | C# | .NET | write protection | Excel workbook password | protect document properties | author metadata | read‑only recommendation | prevent metadata editing | Workbook.Settings.WriteProtection
// Common Searches: Aspose.Cells set workbook author protection | C# write protect Excel file metadata | How to password protect Excel workbook properties with Aspose | Recommend read‑only mode Aspose.Cells | Prevent editing of Excel file properties .NET
// Developer Intent: Apply write‑protection to an Excel workbook so its document properties (author, title, etc.) cannot be altered without a password.
// Use Cases: Distribute a template that must retain the original author information across multiple users. | Share confidential financial reports that open as read‑only and block changes to metadata. | Enforce password‑protected, read‑only mode for spreadsheets containing sensitive corporate data.
// AI Prompts: Provide C# code using Aspose.Cells to set write‑protection on workbook properties with a password and read‑only recommendation. | Generate a snippet that saves an Excel file, reloads it, and prints the author, IsWriteProtected, and RecommendReadOnly values. | Explain how Aspose.Cells prevents modifications to default file properties such as author and title.

using System;
using Aspose.Cells;

namespace AsposeCellsWorkbookProtectionDemo
{
    // Shows how to configure Aspose.Cells Workbook.Settings.WriteProtection to set an author, assign a password, recommend read‑only mode, save the file, and verify the protection flags in C#.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Set write‑protection options to prevent editing of document properties
            workbook.Settings.WriteProtection.Author = "John Doe";          // Author of the protected workbook
            workbook.Settings.WriteProtection.Password = "owner123";       // Password required to modify the file
            workbook.Settings.WriteProtection.RecommendReadOnly = true;   // Recommend opening as read‑only

            // Save the protected workbook
            string outputPath = "ProtectedWorkbook.xlsx";
            workbook.Save(outputPath);

            // Load the workbook back to verify the protection settings
            Workbook loadedWorkbook = new Workbook(outputPath);
            Console.WriteLine("Author: " + loadedWorkbook.Settings.WriteProtection.Author);
            Console.WriteLine("Is Write Protected: " + loadedWorkbook.Settings.WriteProtection.IsWriteProtected);
            Console.WriteLine("Read‑Only Recommended: " + loadedWorkbook.Settings.WriteProtection.RecommendReadOnly);
        }
    }
}
