// Title: Create a new Excel workbook with three named worksheets (Finance, HR, Legal) using Aspose.Cells for .NET
// Common Searches: how to add and rename worksheets in an Excel file with Aspose.Cells C# | Aspose.Cells create workbook with multiple sheets programmatically | C# add new worksheet named HR using Aspose.Cells | save Excel workbook as .xlsx after adding custom worksheets Aspose.Cells | example for creating Finance, HR, and Legal worksheets with Aspose.Cells
// Tags: add worksheets Aspose.Cells C# | rename worksheet Aspose.Cells .NET | create workbook multiple sheets Aspose.Cells | save Excel .xlsx Aspose.Cells | programmatic worksheet management Aspose.Cells

using System;
using Aspose.Cells;

// The sample creates a new Excel workbook, renames the default sheet to "Finance", adds two additional worksheets named "HR" and "Legal", and saves the file as "MultipleSignatureLines.xlsx" using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // -------------------- Worksheet 1 --------------------
            // Use the default first worksheet and rename it
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Finance";

            // -------------------- Worksheet 2 --------------------
            // Add a new worksheet for HR
            Worksheet sheet2 = workbook.Worksheets.Add("HR");

            // -------------------- Worksheet 3 --------------------
            // Add a new worksheet for Legal
            Worksheet sheet3 = workbook.Worksheets.Add("Legal");

            // Save the workbook
            workbook.Save("MultipleSignatureLines.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
