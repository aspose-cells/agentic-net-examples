// Title: C# – Create a Write‑Protected Excel Workbook with Author and Password (Aspose.Cells)
// Description: This C# example demonstrates how to generate a new Excel workbook with Aspose.Cells, assign a write‑protection author, set a password that blocks edits, optionally recommend read‑only opening, and save the file as WriteProtectedWorkbook.xlsx.
// Keywords: Aspose.Cells C# write protection | Excel password protection .NET | set write‑protection author Aspose | recommend read‑only workbook | protect workbook from modification | Aspose.Cells workbook settings
// Common Searches: Aspose.Cells set workbook password C# | How to add write protection author to Excel file using Aspose | C# protect Excel workbook read‑only recommendation | Create password‑protected Excel file with Aspose.Cells
// Developer Intent: Apply write protection with a specific author and password to a newly created Excel workbook.
// Use Cases: Distribute a template that users can view but only edit with a password | Secure financial statements while logging the protection creator for compliance | Offer read‑only access to shared reports unless the correct password is entered
// AI Prompts: Generate C# code that protects an existing workbook with a given author, password, and optional read‑only flag using Aspose.Cells. | Explain how to change or remove the write‑protection settings from a workbook created with Aspose.Cells. | Show a reusable method that accepts author, password, and a read‑only option to protect and save a workbook.

using System;
using Aspose.Cells;

// This C# example demonstrates how to generate a new Excel workbook with Aspose.Cells, assign a write‑protection author, set a password that blocks edits, optionally recommend read‑only opening, and save the file as WriteProtectedWorkbook.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Set the author for write protection
        workbook.Settings.WriteProtection.Author = "John Doe";

        // Set the password that protects the workbook from modification
        workbook.Settings.WriteProtection.Password = "password123";

        // (Optional) Recommend opening the file as read‑only
        workbook.Settings.WriteProtection.RecommendReadOnly = true;

        // Save the workbook with the write‑protection settings applied
        workbook.Save("WriteProtectedWorkbook.xlsx");
    }
}
