// Title: Protect only the workbook structure with a password while keeping cells editable on unlocked worksheets using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates an Excel workbook, adds multiple worksheets, writes data to cells, and applies structure‑only protection with a password using Aspose.Cells. | Show how to use Aspose.Cells to prevent sheet reordering, addition, or deletion while allowing all cell contents to remain editable.
// Common Searches: Aspose.Cells C# protect workbook structure without locking cell content | How to enable password‑protected workbook layout only in Aspose.Cells | Prevent sheet reordering in Excel using Aspose.Cells .NET | Structure‑only protection for Excel files while keeping cells editable Aspose.Cells | C# example for password protecting workbook structure with Aspose.Cells
// Tags: workbook structure protection Aspose.Cells | password protect workbook layout C# | allow cell editing on protected workbook | prevent sheet addition deletion Aspose.Cells | Aspose.Cells Protect method usage

using Aspose.Cells;
using System;

// The example creates a new workbook, adds two worksheets, writes sample text to cell A1 on each sheet, then protects only the workbook structure with a password using workbook.Protect(ProtectionType.Structure, "MySecretPassword"), allowing users to edit cell contents while the sheet order and addition/deletion are locked.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and rename it
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Sheet1";

            // Add a second worksheet
            Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");

            // Populate some cells with sample data
            sheet1.Cells["A1"].PutValue("Editable data on Sheet1");
            sheet2.Cells["A1"].PutValue("Editable data on Sheet2");

            // Protect only the workbook structure with a password
            workbook.Protect(ProtectionType.Structure, "MySecretPassword");

            // Save the workbook
            string outputPath = "ProtectedStructure.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
