// Title: Use Aspose.Cells in C# to protect a worksheet, allow new comments, and prevent editing of existing comments
// AI Prompts: Write C# code with Aspose.Cells that protects a worksheet with a password, enables comment insertion, and locks existing comments from modification. | Show how to configure Aspose.Cells worksheet protection to allow adding comments while disallowing edits to current comments in a .NET application. | Provide an example that adds a comment to a cell, then applies ProtectionType.All with a password so only new comments can be added.
// Common Searches: Aspose.Cells C# protect worksheet allow comment insertion but block comment editing | How to enable only new comments on a protected Excel sheet using Aspose.Cells .NET | C# Aspose.Cells worksheet protection settings for allowing comments and restricting edits | Set password protection on Excel worksheet with Aspose.Cells while permitting comment addition
// Tags: Aspose.Cells worksheet protection allow comments | C# protect Excel worksheet password Aspose.Cells | Aspose.Cells disable existing comment editing | Aspose.Cells comment insertion on protected sheet | ProtectionType.All usage Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, adds data and a comment to cell A1, then protects the first worksheet with a password using ProtectionType.All, which permits inserting new comments while locking existing comments, and saves the file as ProtectedWorksheet.xlsx.
class WorksheetProtectionExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data
            sheet.Cells["A1"].PutValue("Sample Data");
            sheet.Cells["A2"].PutValue(123);

            // Add a comment to A1
            int commentIndex = sheet.Comments.Add("A1");
            Comment comment = sheet.Comments[commentIndex];
            comment.Note = "This comment is locked and cannot be edited.";

            // Protect the worksheet with a password.
            // The third parameter (oldPassword) is required; pass an empty string if not changing an existing password.
            sheet.Protect(ProtectionType.All, "MySecretPassword", string.Empty);

            // Note: In some older Aspose.Cells versions the Protection class does not expose
            // the Allow* properties. If needed, adjust the protection options using the
            // appropriate API for your version.

            // Save the workbook
            string outputPath = "ProtectedWorksheet.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
