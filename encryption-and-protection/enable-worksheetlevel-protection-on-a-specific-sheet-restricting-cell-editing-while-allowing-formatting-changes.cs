// Title: Protect an Excel worksheet with Aspose.Cells .NET – block edits, allow formatting
// Description: Shows how to use Aspose.Cells for .NET to protect a specific worksheet, disable editing of locked cells, permit cell formatting, apply a password, and save the workbook. The sample also creates the output folder and handles exceptions.
// Keywords: Aspose.Cells | worksheet protection | .NET | C# | Excel sheet password | allow cell formatting | disable cell editing | Protect method | ProtectionType.All | Excel security
// Common Searches: Aspose.Cells protect worksheet allow formatting | C# protect Excel sheet with password Aspose.Cells | disable cell editing but keep formatting Aspose.Cells | worksheet protection options .NET Aspose | AllowFormattingCell example Aspose.Cells
// Developer Intent: Enable worksheet protection that blocks content changes while allowing formatting modifications.
// Use Cases: Distribute a template where users can style cells but cannot alter underlying data. | Lock calculated values in a financial report while still permitting column width or color adjustments. | Automate export of a protected workbook with a known password to preserve data integrity and visual flexibility.
// AI Prompts: Provide C# code using Aspose.Cells to protect a worksheet with a password, prevent content edits, and allow formatting changes. | Show an example of setting worksheet protection options in Aspose.Cells so only cell styling is permitted. | Explain how to enable worksheet‑level protection in Aspose.Cells for .NET while keeping the sheet editable for formatting.

using System;
using System.IO;
using Aspose.Cells;

// Shows how to use Aspose.Cells for .NET to protect a specific worksheet, disable editing of locked cells, permit cell formatting, apply a password, and save the workbook. The sample also creates the output folder and handles exceptions.
public class WorksheetProtectionDemo
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
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Configure protection settings
        Protection protection = sheet.Protection;
        protection.AllowEditingContent = false;      // Disallow editing of locked cells
        protection.AllowFormattingCell = true;       // Allow cell formatting while protected

        // Apply protection with password
        sheet.Protect(ProtectionType.All, "pwd123", null);

        // Determine output file path
        string outputPath = "WorksheetProtected.xlsx";

        // Ensure the output directory exists
        string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
        if (!Directory.Exists(directory))
        {
            Directory.CreateDirectory(directory);
        }

        // Save the workbook
        workbook.Save(outputPath);
        Console.WriteLine($"Workbook saved to {outputPath}");
    }
}
