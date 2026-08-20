// Title: Aspose.Cells .NET – Protect Workbook Structure to Block Adding Sheets While Allowing Rename
// Description: Demonstrates how to create a workbook, add worksheets, and apply structure protection with a password using Aspose.Cells for .NET. The protection prevents inserting, deleting, or moving sheets but still permits renaming existing worksheets, then saves the file as an .xlsx document.
// Keywords: Aspose.Cells protect structure | C# workbook protection | prevent adding worksheets | allow sheet rename | Excel structure password .NET
// Common Searches: Aspose.Cells protect workbook structure C# | stop users adding new sheets but allow rename Aspose.Cells | set password for Excel workbook structure protection .NET | how to lock sheet order with Aspose.Cells
// Developer Intent: Apply password‑protected structure protection so new worksheets cannot be added while existing sheets remain rename‑able.
// Use Cases: Distribute a template where the sheet layout must stay fixed but users can label sheets for clarity. | Secure a financial model to prevent accidental sheet insertion while allowing custom names. | Provide a reporting workbook that maintains order yet supports personalized sheet titles.
// AI Prompts: Show code to remove structure protection from a workbook using Aspose.Cells. | Explain how to protect only the workbook windows without affecting the structure in C#. | Give an example of changing the protection password of an already saved Excel file with Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, add worksheets, and apply structure protection with a password using Aspose.Cells for .NET. The protection prevents inserting, deleting, or moving sheets but still permits renaming existing worksheets, then saves the file as an .xlsx document.
    public class WorkbookStructureProtectionDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and add a couple of worksheets
                Workbook workbook = new Workbook();
                workbook.Worksheets.Add("DataSheet");
                workbook.Worksheets.Add("Summary");

                // Protect the workbook structure (prevents adding, deleting, moving worksheets)
                // Renaming existing worksheets remains allowed.
                // Use a password so the protection can be removed later if needed.
                workbook.Protect(ProtectionType.Structure, "mySecretPwd");

                // Save the protected workbook
                workbook.Save("WorkbookWithStructureProtection.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            WorkbookStructureProtectionDemo.Run();
        }
    }
}
