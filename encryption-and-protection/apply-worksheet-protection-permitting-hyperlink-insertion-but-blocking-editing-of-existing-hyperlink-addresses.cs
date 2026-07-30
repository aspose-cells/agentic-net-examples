// Title: Aspose.Cells C# – Protect Worksheet, Permit Hyperlink Insertion, Block Existing Link Editing
// Description: Demonstrates how to protect an Excel worksheet with a password, enable users to add new hyperlinks, and prevent any changes to existing hyperlink URLs using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# worksheet protection | .NET Excel security | allow hyperlink insertion | prevent hyperlink edit | password protected sheet | ProtectionType.All | Excel hyperlink protection | global Excel security
// Common Searches: Aspose.Cells allow adding hyperlinks on protected sheet | C# protect worksheet but keep hyperlink insertion enabled | block editing of existing hyperlink URLs in Excel with Aspose | set worksheet password and restrict content editing Aspose.Cells
// Developer Intent: Protect a worksheet so collaborators can insert new hyperlinks while the original hyperlink addresses remain immutable.
// Use Cases: Shared template where source URLs are locked but team members can add reference links. | Financial report that safeguards vendor links yet lets analysts attach supplemental resources. | Internal knowledge base workbook that preserves approved documentation URLs while permitting staff to add new citations.
// AI Prompts: Generate C# code using Aspose.Cells to protect a sheet, enable hyperlink insertion, and disable editing of existing hyperlink addresses. | Explain how to modify the protection settings to also allow editing of cell comments while keeping hyperlink insertion allowed. | Provide a guide for applying the same hyperlink‑insertion protection across multiple worksheets in a workbook with Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to protect an Excel worksheet with a password, enable users to add new hyperlinks, and prevent any changes to existing hyperlink URLs using Aspose.Cells for .NET.
    public class WorksheetProtectionHyperlinkDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add an initial hyperlink that will be protected from editing
                worksheet.Hyperlinks.Add("A1", 1, 1, "https://original.com");

                // Access the worksheet protection settings
                Protection protection = worksheet.Protection;

                // Allow users to insert new hyperlinks while the sheet is protected
                protection.AllowInsertingHyperlink = true;

                // Disallow editing of cell contents (including existing hyperlink addresses)
                protection.AllowEditingContent = false;

                // Set a password for the protection
                protection.Password = "securePwd";

                // Apply protection to all protection types
                worksheet.Protect(ProtectionType.All);

                // Save the workbook
                workbook.Save("WorksheetProtectionHyperlinkDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            WorksheetProtectionHyperlinkDemo.Run();
        }
    }
}
