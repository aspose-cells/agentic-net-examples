// Title: Aspose.Cells C# – Protect Worksheet: Block Column Insertion, Allow Column Resizing
// Description: Shows how to use Aspose.Cells for .NET to protect a worksheet, prevent users from inserting columns, enable column width adjustments, apply a password, and save the workbook as an XLSX file.
// Keywords: Aspose.Cells worksheet protection C# | disable column insertion Aspose.Cells | allow column resizing Aspose.Cells | worksheet password protection .NET | protect worksheet structure Aspose.Cells
// Common Searches: prevent column insertion in Aspose.Cells worksheet | allow column width changes while worksheet is protected Aspose.Cells | set password for worksheet protection C# Aspose.Cells | Aspose.Cells protect sheet but enable formatting | how to lock worksheet structure Aspose.Cells .NET
// Developer Intent: Protect a worksheet so end users can resize columns but cannot add new columns.
// Use Cases: Template where layout must stay fixed while users adjust column widths for readability. | Shared report that requires a password to lock structure yet permits personal column formatting. | Data export that maintains column integrity but allows visual customization of column size.
// AI Prompts: Provide C# code with Aspose.Cells to protect a worksheet, disable column insertion, enable column resizing, and set a password. | Explain the effect of AllowInsertingColumn and AllowFormattingColumn when using ProtectionType.All in Aspose.Cells .NET. | Show an example of protecting a worksheet while allowing only column width changes for end users.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Shows how to use Aspose.Cells for .NET to protect a worksheet, prevent users from inserting columns, enable column width adjustments, apply a password, and save the workbook as an XLSX file.
    public class WorksheetProtectionDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Access the protection settings of the worksheet
                Protection protection = worksheet.Protection;

                // Disallow insertion of columns while the sheet is protected
                protection.AllowInsertingColumn = false;

                // Allow users to resize (format) columns
                protection.AllowFormattingColumn = true;

                // Set a password for the protection
                protection.Password = "password123";

                // Apply protection with all protection types
                worksheet.Protect(ProtectionType.All);

                // Define output file path
                string outputPath = "WorksheetProtection.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            WorksheetProtectionDemo.Run();
        }
    }
}
