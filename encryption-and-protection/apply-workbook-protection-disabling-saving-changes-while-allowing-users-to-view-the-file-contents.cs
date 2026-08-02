// Title: Aspose.Cells .NET – Apply owner password and enforce read‑only mode for Excel workbooks
// Description: Shows how to create a Workbook, assign an owner password through WriteProtection, enable the RecommendReadOnly flag, and save the file so it opens as view‑only, blocking users from saving changes.
// Keywords: Aspose.Cells | C# write protection | owner password Excel | read only workbook .NET | RecommendReadOnly | prevent saving Excel file | Excel file protection C# | Aspose.Cells encryption
// Common Searches: Aspose.Cells set owner password C# | Make Excel file read only with Aspose.Cells | Disable saving changes in an Excel workbook using Aspose.Cells .NET | How to use RecommendReadOnly flag Aspose.Cells | Protect workbook from editing Aspose.Cells
// Developer Intent: The developer wants to safeguard a workbook so users can view its contents but cannot modify or overwrite it, using an owner password and a read‑only recommendation.
// Use Cases: Distribute financial reports that must remain unchanged by recipients. | Provide a template to clients that should stay intact while they work on copies. | Publish internal policy documents as view‑only Excel files to preserve content integrity.
// AI Prompts: Give C# code that sets an owner password and enables RecommendReadOnly with Aspose.Cells. | Explain how to protect an existing workbook from being saved, including changing or removing the password, using Aspose.Cells for .NET. | Show how to toggle the read‑only recommendation on a workbook programmatically.

using Aspose.Cells;

// Shows how to create a Workbook, assign an owner password through WriteProtection, enable the RecommendReadOnly flag, and save the file so it opens as view‑only, blocking users from saving changes.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Set a password that protects the file from being modified
        workbook.Settings.WriteProtection.Password = "ownerPwd";

        // Recommend the file to be opened as read‑only (users can view but not save changes)
        workbook.Settings.WriteProtection.RecommendReadOnly = true;

        // Save the workbook to disk
        workbook.Save("ReadOnlyWorkbook.xlsx");
    }
}
