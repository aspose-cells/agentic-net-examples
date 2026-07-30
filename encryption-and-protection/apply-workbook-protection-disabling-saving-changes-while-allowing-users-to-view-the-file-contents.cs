// Title: Apply Write‑Protection to an Excel Workbook with Aspose.Cells for .NET (view‑only, password required to save)
// Description: Shows how to set a password, author name, and RecommendReadOnly flag via Workbook.Settings.WriteProtection, then save the file so users can view it but must enter the password to modify or save changes.
// Keywords: Aspose.Cells | workbook write protection | Excel read‑only .NET | password protect Excel file C# | Settings.WriteProtection | disable saving changes Excel | C# Aspose.Cells example | protect Excel workbook programmatically
// Common Searches: Aspose.Cells set workbook write protection C# | How to make Excel file read‑only with password using Aspose.Cells | Prevent saving changes in Excel workbook .NET | Aspose.Cells recommend read‑only mode | C# protect Excel workbook from editing
// Developer Intent: Create an Excel workbook that can be opened for viewing but requires a password to edit or save any changes.
// Use Cases: Distribute a template that stakeholders can view but only authorized users can edit and re‑save. | Provide auditors with a read‑only financial report that can be modified only after entering a password. | Export data for partners where the file is viewable in Excel yet protected against unauthorized modifications.
// AI Prompts: Generate C# code using Aspose.Cells to apply write‑protection with a password and enable RecommendReadOnly for a new workbook. | Explain how to remove or change the write‑protection settings of an existing workbook saved with Aspose.Cells. | Show how to detect if a workbook is write‑protected before attempting to save changes in Aspose.Cells for .NET.

using Aspose.Cells;

// Shows how to set a password, author name, and RecommendReadOnly flag via Workbook.Settings.WriteProtection, then save the file so users can view it but must enter the password to modify or save changes.
class WorkbookWriteProtectionDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Configure write‑protection: set a password required to modify the file,
        // specify the author, and recommend opening the file as read‑only.
        workbook.Settings.WriteProtection.Password = "modify123";
        workbook.Settings.WriteProtection.Author = "Jane Doe";
        workbook.Settings.WriteProtection.RecommendReadOnly = true;

        // Save the workbook. Users can view the contents, but saving changes
        // requires the password set above.
        workbook.Save("ReadOnlyProtectedWorkbook.xlsx");

        // Release resources
        workbook.Dispose();
    }
}
