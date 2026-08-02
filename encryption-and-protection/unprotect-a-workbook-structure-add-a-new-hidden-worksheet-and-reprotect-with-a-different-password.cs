// Title: Aspose.Cells .NET: Unprotect, add hidden sheet, and re‑protect workbook with new password
// Description: Demonstrates how to remove structure protection from a Workbook, insert a hidden worksheet, and apply a new password to the workbook structure using Aspose.Cells for .NET, then save the file.
// Keywords: Aspose.Cells unprotect workbook | add hidden worksheet C# | protect workbook structure password | change workbook protection Aspose | hide sheet programmatically .NET
// Common Searches: unprotect workbook structure Aspose.Cells C# | add hidden sheet and protect workbook with new password | change workbook protection password after editing sheets | Aspose.Cells hide worksheet and re‑apply protection | C# code to modify protected Excel file
// Developer Intent: Remove existing workbook structure protection, insert a hidden worksheet, then protect the workbook again using a different password.
// Use Cases: Create a template with a concealed configuration sheet that is secured with a new password before distribution. | Update a protected workbook by adding an audit sheet while rotating the protection password for compliance. | Generate reports that contain internal data on a hidden sheet, then lock the workbook for end‑user access.
// AI Prompts: Write C# code with Aspose.Cells to unprotect a workbook, add a hidden worksheet, and protect it with a new password. | Explain how to change the workbook structure protection password after modifying worksheets using Aspose.Cells for .NET. | Suggest robust error‑handling patterns when unprotecting and re‑protecting an Excel workbook with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Demonstrates how to remove structure protection from a Workbook, insert a hidden worksheet, and apply a new password to the workbook structure using Aspose.Cells for .NET, then save the file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Protect the workbook structure with an initial password
                workbook.Protect(ProtectionType.Structure, "oldPassword");

                // Unprotect the workbook using the same password
                workbook.Unprotect("oldPassword");

                // Add a new worksheet
                int newSheetIndex = workbook.Worksheets.Add();
                Worksheet hiddenSheet = workbook.Worksheets[newSheetIndex];

                // Hide the newly added worksheet (false makes the sheet hidden)
                hiddenSheet.IsVisible = false;

                // Re‑protect the workbook structure with a different password
                workbook.Protect(ProtectionType.Structure, "newPassword");

                // Save the workbook
                string outputPath = "ProtectedWorkbook.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
