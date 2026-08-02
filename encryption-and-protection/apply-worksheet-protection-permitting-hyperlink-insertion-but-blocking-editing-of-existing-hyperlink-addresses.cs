// Title: Protect an Aspose.Cells worksheet to allow new hyperlinks while preventing edits to existing links
// Description: Shows how to secure a worksheet with a password, enable insertion of fresh hyperlinks (AllowInsertingHyperlink = true), and block modifications to existing hyperlink addresses (AllowEditingContent = false) using Aspose.Cells for .NET. The sample adds a link to A1 and saves the file.
// Keywords: Aspose.Cells | C# worksheet protection | allow hyperlink insertion | prevent hyperlink editing | Protection.AllowInsertingHyperlink | Protection.AllowEditingContent | .NET Excel security | password protected sheet | hyperlink management
// Common Searches: Aspose.Cells allow hyperlink insertion on protected sheet | prevent editing of existing hyperlinks Aspose.Cells | worksheet protection password C# Aspose.Cells | enable hyperlink addition while sheet is locked Aspose.Cells | protect Excel sheet but allow new links using Aspose.Cells
// Developer Intent: The developer needs a protected worksheet where users can add new hyperlinks but cannot change the URLs of hyperlinks that are already present.
// Use Cases: Distribute a template that lets collaborators add reference links without altering predefined navigation URLs. | Create a shared report where contributors can insert citations while the original hyperlinks remain immutable. | Build an internal dashboard that permits linking to new resources but safeguards existing link destinations.
// AI Prompts: Write C# code with Aspose.Cells to protect a worksheet, enable new hyperlink insertion, and block edits to existing hyperlink addresses. | Show how to configure Worksheet.Protection in Aspose.Cells for .NET to allow adding hyperlinks while disabling content changes, then save the workbook. | Provide an example that adds a hyperlink after setting AllowInsertingHyperlink = true and AllowEditingContent = false on a password‑protected sheet.

using System;
using Aspose.Cells;

// Shows how to secure a worksheet with a password, enable insertion of fresh hyperlinks (AllowInsertingHyperlink = true), and block modifications to existing hyperlink addresses (AllowEditingContent = false) using Aspose.Cells for .NET. The sample adds a link to A1 and saves the file.
public class WorksheetProtectionHyperlinkDemo
{
    public static void Main(string[] args)
    {
        try
        {
            Run();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    public static void Run()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Access the worksheet protection settings
        Protection protection = sheet.Protection;

        // Allow insertion of new hyperlinks while the sheet is protected
        protection.AllowInsertingHyperlink = true;

        // Disallow editing of cell contents (prevents changing existing hyperlink addresses)
        protection.AllowEditingContent = false;

        // Set a password for the protection
        protection.Password = "pwd123";

        // Apply protection to the worksheet (protect all aspects)
        sheet.Protect(ProtectionType.All);

        // Add a hyperlink to demonstrate that insertion is allowed
        sheet.Hyperlinks.Add("A1", 1, 1, "https://www.example.com");

        // Save the workbook
        string outputPath = "WorksheetProtectionHyperlink.xlsx";
        workbook.Save(outputPath);
        Console.WriteLine($"Workbook saved to {outputPath}");
    }
}
