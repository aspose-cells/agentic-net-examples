// Title: Save an Aspose.Cells workbook containing watch‑window style SUM and AVERAGE formulas to an XLSX file in C#
// Common Searches: aspnet cells c# save workbook with formulas to xlsx | how to export watch window calculations to Excel using Aspose.Cells | C# Aspose.Cells create workbook add SUM formula and save | export workbook after assigning cell formulas with Aspose.Cells .NET | Aspose.Cells example for watch‑window style formula export
// Tags: Aspose.Cells output Xlsx file | watch window formula configuration Aspose.Cells | C# add SUM and AVERAGE formulas Aspose.Cells | generate Excel file with formulas Aspose.Cells | set cell formula before exporting Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a new Workbook, places SUM(A1:A10) in B1 and AVERAGE(B1:B10) in C1 to simulate a watch window, then writes the workbook to 'WatchWindowDemo.xlsx' using the Xlsx SaveFormat, and prints the full file path while handling any exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add formulas to cells B1 and C1 (watch-like behavior)
            // Sum of A1:A10 displayed at B1
            sheet.Cells["B1"].Formula = "=SUM(A1:A10)";
            // Average of B1:B10 displayed at C1
            sheet.Cells["C1"].Formula = "=AVERAGE(B1:B10)";

            // Define output file path
            string outputPath = "WatchWindowDemo.xlsx";

            // Save the workbook to an Excel file
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
