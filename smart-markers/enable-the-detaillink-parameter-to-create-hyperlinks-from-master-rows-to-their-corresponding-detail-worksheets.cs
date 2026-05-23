using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;   // For PivotTable and PivotFieldType enums

public class DetailLinkDemo
{
    public static void Main()
    {
        try
        {
            Run();
            Console.WriteLine("Workbook created successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    public static void Run()
    {
        // Create a new workbook and get the first worksheet (master sheet)
        Workbook workbook = new Workbook();
        Worksheet masterSheet = workbook.Worksheets[0];
        masterSheet.Name = "Master";

        // Populate master sheet with sample data for a pivot table
        masterSheet.Cells["A1"].PutValue("Category");
        masterSheet.Cells["B1"].PutValue("Amount");
        masterSheet.Cells["A2"].PutValue("A");
        masterSheet.Cells["B2"].PutValue(100);
        masterSheet.Cells["A3"].PutValue("B");
        masterSheet.Cells["B3"].PutValue(200);
        masterSheet.Cells["A4"].PutValue("A");
        masterSheet.Cells["B4"].PutValue(150);
        masterSheet.Cells["A5"].PutValue("B");
        masterSheet.Cells["B5"].PutValue(250);

        // Add a pivot table to the master sheet
        int pivotIndex = masterSheet.PivotTables.Add("A1:B5", "D3", "PivotTable1");
        PivotTable pivot = masterSheet.PivotTables[pivotIndex];
        pivot.AddFieldToArea(PivotFieldType.Row, "Category");
        pivot.AddFieldToArea(PivotFieldType.Data, "Amount");

        // Create a detail worksheet that will receive drill‑down data
        Worksheet detailSheet = workbook.Worksheets.Add("Detail");

        // Enable the built‑in property indicating that hyperlinks are up‑to‑date
        workbook.BuiltInDocumentProperties.LinksUpToDate = true;

        // Add hyperlinks from each pivot row (master) to the detail sheet.
        // Internal hyperlink format uses a leading '#'.
        masterSheet.Hyperlinks.Add("D4", 1, 1, "#Detail!A1");   // Link from first data row
        masterSheet.Hyperlinks.Add("D5", 1, 1, "#Detail!A10"); // Link from second data row

        // Optionally, set display text for the hyperlinks
        masterSheet.Hyperlinks[0].TextToDisplay = "View Detail A";
        masterSheet.Hyperlinks[1].TextToDisplay = "View Detail B";

        // Define output path and ensure the directory exists
        string outputPath = "DetailLinkDemo.xlsx";
        string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
        if (!Directory.Exists(outputDir))
        {
            Directory.CreateDirectory(outputDir);
        }

        // Save the workbook
        workbook.Save(outputPath);
    }
}