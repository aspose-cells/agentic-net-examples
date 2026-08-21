// Title: C# – Apply Write Protection and Recommend Read‑Only for Excel Workbook with Aspose.Cells
// Description: Demonstrates how to create an Excel workbook, set an owner password, enable the RecommendReadOnly flag, and save the file so users can view it but cannot save changes without the password using Aspose.Cells for .NET.
// Keywords: Aspose.Cells write protection C# | Excel recommend read only | password protected workbook .NET | prevent saving changes Excel | Aspose.Cells workbook security
// Common Searches: Aspose.Cells set write protection password | C# make Excel file read only with Aspose | How to recommend read only in Aspose.Cells | Prevent users from saving Excel changes .NET | Enable write protection in Aspose.Cells workbook
// Developer Intent: Protect an Excel workbook so it opens for viewing but cannot be saved without providing the owner password.
// Use Cases: Distribute a template that must stay unchanged unless authorized. | Share a financial report that users can view but not overwrite. | Provide a read‑only version of a document while allowing privileged edits after password entry.
// AI Prompts: Write C# code with Aspose.Cells to add write protection, set an owner password, and enable RecommendReadOnly. | Explain how to modify the sample so the workbook becomes editable after entering the correct password. | List steps to test that the saved workbook opens in read‑only mode when opened without a password.

using System;
using Aspose.Cells;

namespace AsposeCellsWriteProtectionDemo
{
    // Demonstrates how to create an Excel workbook, set an owner password, enable the RecommendReadOnly flag, and save the file so users can view it but cannot save changes without the password using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add some sample data (optional)
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Read‑only demo");

            // Access write‑protection settings
            WriteProtection writeProtection = workbook.Settings.WriteProtection;

            // Set a password required to modify the file
            writeProtection.Password = "ownerPassword";

            // Recommend the file to be opened as read‑only
            writeProtection.RecommendReadOnly = true;

            // Save the workbook; users can view it but cannot save changes without the password
            string outputPath = "ReadOnlyProtectedWorkbook.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            // Clean up
            workbook.Dispose();

            Console.WriteLine($"Workbook saved to '{outputPath}' with write protection.");
        }
    }
}
