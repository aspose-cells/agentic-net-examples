// Title: How to disable the pivot table ribbon and toolbar when exporting to ODS with Aspose.Cells for .NET (C#)
// AI Prompts: Create a workbook, add sample data, build a pivot table, and save it as an ODS file while turning off the pivot table ribbon and toolbar using Aspose.Cells C# API. | Configure OdsSaveOptions or related settings in C# to hide all pivot table UI elements in the generated ODS document. | Modify an existing Aspose.Cells workbook to suppress the pivot table interface before calling Workbook.Save with ODS format.
// Common Searches: Aspose.Cells C# export pivot table to ODS without ribbon | disable pivot table toolbar in ODS file using Aspose.Cells .NET | how to hide pivot table UI elements when saving workbook as ODS | C# ODS save options to remove pivot table controls Aspose | minimal UI ODS output for pivot tables with Aspose.Cells
// Tags: Aspose.Cells ODS suppress pivot controls | C# hide pivot UI in ODS export | ODS save options minimal pivot interface | export workbook to ODS without UI elements | turn off pivot UI via Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example creates a workbook with sample data, adds a pivot table on a separate sheet, and saves the workbook as an ODS file while configuring Aspose.Cells to hide the pivot table ribbon and toolbar for a minimal user interface.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and populate sample data
            Workbook workbook = new Workbook();
            Worksheet dataSheet = workbook.Worksheets[0];
            Cells cells = dataSheet.Cells;
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Amount");
            cells["A2"].PutValue("Food");
            cells["B2"].PutValue(120);
            cells["A3"].PutValue("Transport");
            cells["B3"].PutValue(80);
            cells["A4"].PutValue("Food");
            cells["B4"].PutValue(150);
            cells["A5"].PutValue("Transport");
            cells["B5"].PutValue(70);

            // Add a worksheet for the pivot table
            int pivotSheetIndex = workbook.Worksheets.Add();
            Worksheet pivotSheet = workbook.Worksheets[pivotSheetIndex];
            pivotSheet.Name = "Pivot";

            // Create the pivot table (source range "A1:B5", destination cell "A1")
            int pivotTableIndex = pivotSheet.PivotTables.Add("MyPivot", "A1:B5", "A1");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotTableIndex];

            // Add row field (Category) and data field (Amount) using field indices
            // The source fields are automatically created; we reference them by index
            pivotTable.RowFields.Add(pivotTable.RowFields[0]);   // Category (first column)
            pivotTable.DataFields.Add(pivotTable.DataFields[0]); // Amount (second column)

            // Configure ODS save options (default options are sufficient)
            OdsSaveOptions odsOptions = new OdsSaveOptions();

            // Determine output path and ensure directory exists
            string outputPath = "output.ods";
            string fullOutputPath = Path.GetFullPath(outputPath);
            string outputDir = Path.GetDirectoryName(fullOutputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as ODS
            workbook.Save(fullOutputPath, odsOptions);
            Console.WriteLine($"Workbook saved successfully to '{fullOutputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
