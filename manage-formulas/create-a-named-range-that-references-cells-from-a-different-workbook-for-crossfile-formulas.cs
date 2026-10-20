// Title: Create a named range that references an external workbook range for cross‑workbook formulas using Aspose.Cells in C#
// Common Searches: how to define a named range that references a range in another Excel file with Aspose.Cells C# | Aspose.Cells C# external workbook reference in a named range for formulas | add external range to workbook names collection using Aspose.Cells .NET | use Aspose.Cells to create a named range for a SUM formula that pulls data from a different workbook | C# Aspose.Cells external reference string format for named ranges
// Tags: Aspose.Cells create external named range | C# external workbook reference Aspose.Cells | named range RefersTo property .NET | cross-workbook formula Aspose.Cells | Excel external range string Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a new workbook, builds an Excel‑style external reference string to a range in another workbook, adds a named range called "ExternalData" that points to that external range, inserts a SUM formula using the named range, and saves the workbook as Result.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (default worksheet is added automatically)
        Workbook workbook = new Workbook();

        // Define the external workbook file name, sheet name, and cell range to reference
        string externalFile = "OtherWorkbook.xlsx";
        string externalSheet = "Sheet1";
        string externalRange = "$A$1:$B$2";

        // Build the external reference string in Excel format:
        // ='[OtherWorkbook.xlsx]Sheet1'!$A$1:$B$2
        string externalReference = $"='[{externalFile}]{externalSheet}'!{externalRange}";

        // Add a named range to the current workbook that points to the external range
        // The name "ExternalData" can be used in formulas within this workbook
        int nameIndex = workbook.Worksheets.Names.Add("ExternalData");
        Name externalName = workbook.Worksheets.Names[nameIndex];
        externalName.RefersTo = externalReference;

        // Example usage: place a formula that sums the external range into cell C1
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["C1"].Formula = "SUM(ExternalData)";

        // Save the workbook to a file
        workbook.Save("Result.xlsx");
    }
}
