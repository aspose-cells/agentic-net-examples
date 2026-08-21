// Title: Password‑protect rows 20‑25 in an Excel sheet using Aspose.Cells for .NET (EntireRow)
// Description: Creates a workbook, fills cells A1:E30, selects rows 20‑25 with the EntireRow property, adds them to the AllowEditRanges collection, assigns a password, protects the worksheet with all protection types, and saves the file as ProtectedRows20to25.xlsx.
// Keywords: Aspose.Cells | .NET | C# | protect rows | EntireRow property | password protection | AllowEditRanges | Excel worksheet protection | range protection | specific rows
// Common Searches: Aspose.Cells protect specific rows | How to password protect rows in Excel using Aspose.Cells | EntireRow property example Aspose.Cells | C# protect rows 20 to 25 Aspose.Cells | Set AllowEditRanges password Aspose.Cells
// Developer Intent: The developer needs to lock rows 20‑25 of a worksheet with a password by using the EntireRow property in Aspose.Cells for .NET.
// Use Cases: Lock header or template rows in a generated report while keeping data rows editable. | Secure confidential financial rows in a shared workbook to prevent accidental changes. | Create a read‑only section in a data‑entry form that users cannot modify.
// AI Prompts: Generate C# code that password‑protects rows 10‑15 in an Excel file using Aspose.Cells. | Explain how to use AllowEditRanges with the EntireRow property to lock specific rows in Aspose.Cells for .NET. | Show the steps to unprotect a password‑protected row range in an Aspose.Cells workbook.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a workbook, fills cells A1:E30, selects rows 20‑25 with the EntireRow property, adds them to the AllowEditRanges collection, assigns a password, protects the worksheet with all protection types, and saves the file as ProtectedRows20to25.xlsx.
    public class ProtectRowsWithEntireRowDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data
                for (int row = 0; row < 30; row++)
                {
                    for (int col = 0; col < 5; col++)
                    {
                        worksheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                    }
                }

                // Create a range covering rows 20‑25 and get the entire rows
                Aspose.Cells.Range rowsRange = worksheet.Cells.CreateRange("A20:A25").EntireRow;

                // Add a protected range for those rows
                ProtectedRangeCollection allowEditRanges = worksheet.AllowEditRanges;
                int protectedIndex = allowEditRanges.Add(
                    "Rows20to25",
                    rowsRange.FirstRow,
                    rowsRange.FirstColumn,
                    rowsRange.FirstRow + rowsRange.RowCount - 1,
                    rowsRange.FirstColumn + rowsRange.ColumnCount - 1);

                // Set a password for the protected range
                ProtectedRange protectedRange = allowEditRanges[protectedIndex];
                protectedRange.Password = "MySecretPassword";

                // Protect the worksheet (all protection types)
                worksheet.Protect(ProtectionType.All);

                // Save the workbook
                string outputPath = "ProtectedRows20to25.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            ProtectRowsWithEntireRowDemo.Run();
        }
    }
}
