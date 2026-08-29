// Title: Add a line sparkline to cells A1:D1, place it in E1, and save the workbook as XLSX using Aspose.Cells for .NET
// AI Prompts: Write C# code that creates a new workbook, populates A1:D1 with values, adds a line‑type sparkline referencing that range, positions the sparkline in cell E1, and saves the file to a given XLSX path with Aspose.Cells. | Demonstrate how to enable markers and adjust other SparklineGroup settings before calling Workbook.Save to generate an XLSX file containing the sparkline using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# add line sparkline to a specific cell range and save as .xlsx | How to programmatically create sparkline groups in Excel with Aspose.Cells .NET | Saving an Excel workbook that contains sparkline data using Aspose.Cells C# example | Customize sparkline markers with Aspose.Cells before saving the workbook
// Tags: create line sparkline Aspose.Cells C# | save workbook as xlsx with sparkline Aspose.Cells | sparklinegroup customization Aspose.Cells | add sparkline to cell range Aspose.Cells | Aspose.Cells generate sparkline programmatically

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new workbook, writes sample numbers to A1:D1, adds a line‑type sparkline that references this range and is placed in cell E1, optionally configures the SparklineGroup (e.g., shows markers), and saves the result as an XLSX file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data that the sparkline will represent
        sheet.Cells["A1"].PutValue(5);
        sheet.Cells["B1"].PutValue(2);
        sheet.Cells["C1"].PutValue(1);
        sheet.Cells["D1"].PutValue(3);

        // Define the cell where the sparkline will be placed (E1)
        CellArea sparkArea = new CellArea
        {
            StartRow = 0,
            EndRow = 0,
            StartColumn = 4,
            EndColumn = 4
        };

        // Add a line‑type sparkline group that uses the data range A1:D1
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, sheet.Name + "!A1:D1", false, sparkArea);
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // (Optional) Customize sparkline appearance here, e.g.:
        // group.ShowMarkers = true;

        // Save the workbook with the new sparkline to an XLSX file
        string outputPath = @"C:\Temp\SparklineWorkbook.xlsx";
        workbook.Save(outputPath); // Uses Workbook.Save(string) rule
    }
}
