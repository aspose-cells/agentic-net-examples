// Title: Write‑protect an Excel workbook with author, password and read‑only recommendation using Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, assign an audit author, set a protection password, enable a read‑only recommendation, save the file, reload it, and read back the protection properties with Aspose.Cells for C#.
// Keywords: Aspose.Cells | write protection | Excel author metadata | password protected workbook | recommend read only | C# | .NET | audit tracking | global compliance | protect Excel file
// Common Searches: Aspose.Cells set workbook author for write protection | C# protect Excel file with password and read‑only flag | verify write protection settings after saving workbook | add audit information to protected Excel workbook | how to enable recommend read only in Aspose.Cells
// Developer Intent: Enable write‑protection on an Excel workbook, record the protecting team as the author for audit purposes, require a password for edits, and suggest opening the file as read‑only.
// Use Cases: Distribute financial or regulatory reports that must stay unchanged unless an authorized user supplies a password. | Create template files for internal teams where the protection author logs responsibility and compliance. | Implement audit‑ready Excel documents that automatically expose the author and protection status when opened.
// AI Prompts: Write C# code with Aspose.Cells to apply write‑protection, set an author, add a password, enable read‑only recommendation, and then display the protection details. | Explain the purpose of WriteProtection.Author, Password, RecommendReadOnly, and IsWriteProtected in Aspose.Cells and how to read them after loading a workbook. | Show how to modify the author and password of an already protected workbook using Aspose.Cells without losing existing data.

using System;
using Aspose.Cells;

namespace AsposeCellsWriteProtectionDemo
{
    // Demonstrates how to create a workbook, assign an audit author, set a protection password, enable a read‑only recommendation, save the file, reload it, and read back the protection properties with Aspose.Cells for C#.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Set write‑protection options
            // Author for audit tracking
            workbook.Settings.WriteProtection.Author = "Audit Team";

            // Password required to modify the workbook
            workbook.Settings.WriteProtection.Password = "SecurePass123";

            // Recommend opening the file as read‑only
            workbook.Settings.WriteProtection.RecommendReadOnly = true;

            // Save the protected workbook
            string outputPath = "WriteProtectedWorkbook.xlsx";
            workbook.Save(outputPath);

            // Load the saved workbook to verify protection settings
            Workbook loadedWorkbook = new Workbook(outputPath);

            // Output verification information
            Console.WriteLine("Author: " + loadedWorkbook.Settings.WriteProtection.Author);
            Console.WriteLine("Is Write Protected: " + loadedWorkbook.Settings.WriteProtection.IsWriteProtected);
            Console.WriteLine("Read‑Only Recommended: " + loadedWorkbook.Settings.WriteProtection.RecommendReadOnly);
        }
    }
}
