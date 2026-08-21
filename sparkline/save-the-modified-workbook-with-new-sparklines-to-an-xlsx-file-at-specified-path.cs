// Title: Create and Save an XLSX Workbook with a Line Sparkline using Aspose.Cells for .NET
// Description: C# code that creates a Workbook, fills cells A1‑D1 with sample values, inserts a line sparkline into cell E1, turns on markers, and saves the result to a user‑specified path in XLSX format via Aspose.Cells.
// Keywords: Aspose.Cells | C# | sparkline | line sparkline | Workbook | save XLSX | Excel sparkline programmatically | SparklineGroup | add sparkline | export to XLSX
// Common Searches: Aspose.Cells add line sparkline C# | save workbook with sparkline Aspose .NET | create sparkline programmatically Aspose.Cells | export sparkline to XLSX using C# | how to use SparklineGroup in Aspose.Cells
// Developer Intent: Generate an XLSX file that contains a line sparkline placed in a specific cell.
// Use Cases: Build a sales‑trend report where each row’s data is visualized with a line sparkline and the file is distributed to stakeholders. | Automate a daily financial dashboard that adds sparklines to key metrics and exports the workbook for archiving. | Create a utility that inserts sparklines into existing worksheets, customizes markers, and overwrites the original file with the updated version.
// AI Prompts: Write C# code that adds a column sparkline for every data row and saves the workbook as XLSX using Aspose.Cells. | Show how to change sparkline colors, line thickness, and marker styles before exporting the file with Aspose.Cells for .NET. | Provide an example that creates multiple SparklineGroup objects on different ranges and saves the workbook to a user‑provided path.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// C# code that creates a Workbook, fills cells A1‑D1 with sample values, inserts a line sparkline into cell E1, turns on markers, and saves the result to a user‑specified path in XLSX format via Aspose.Cells.
class SparklineWorkbookExample
{
    // Creates a workbook, adds a sparkline, and saves it as XLSX.
    public static void CreateWorkbookWithSparkline(string outputPath)
    {
        // Initialize a new workbook.
        Workbook workbook = new Workbook();

        // Access the first worksheet.
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the sparkline (row 1, columns A‑D).
        sheet.Cells["A1"].PutValue(5);
        sheet.Cells["B1"].PutValue(2);
        sheet.Cells["C1"].PutValue(1);
        sheet.Cells["D1"].PutValue(3);

        // Define the cell area where the sparkline will be placed (cell E1).
        CellArea sparklineLocation = new CellArea
        {
            StartRow = 0,
            EndRow = 0,
            StartColumn = 4,
            EndColumn = 4
        };

        // Add a sparkline group of type Line using the data range A1:D1.
        int groupIndex = sheet.SparklineGroups.Add(
            SparklineType.Line,
            sheet.Name + "!A1:D1",
            false,
            sparklineLocation);

        // Optional: customize the sparkline group (e.g., show markers).
        SparklineGroup group = sheet.SparklineGroups[groupIndex];
        group.ShowMarkers = true;

        // Save the workbook to the specified path in XLSX format.
        workbook.Save(outputPath, SaveFormat.Xlsx);
    }

    // Example entry point.
    static void Main()
    {
        string outputFile = "SparklineWorkbook.xlsx";
        CreateWorkbookWithSparkline(outputFile);
        Console.WriteLine($"Workbook saved to {outputFile}");
    }
}
