using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Populate sample data with a header row
            sheet.Cells["A1"].PutValue("Item");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("Orange");
            sheet.Cells["B3"].PutValue(15);
            sheet.Cells["A4"].PutValue("Banana");
            sheet.Cells["B4"].PutValue(20);
            // Additional rows can be added later; the named range will adjust automatically

            // Create a named range that expands dynamically based on the number of rows in column A
            // OFFSET(start, rows, cols, height, width)
            // start = A2 (first data row), height = COUNTA(A:A)-1 (exclude header)
            int nameIdx = workbook.Worksheets.Names.Add("DynamicItems");
            Name dynamicName = workbook.Worksheets.Names[nameIdx];
            dynamicName.RefersTo = $"=OFFSET({sheet.Name}!$A$2,0,0,COUNTA({sheet.Name}!$A:$A)-1,1)";

            // Retrieve the range to verify the address (optional)
            Aspose.Cells.Range dynRange = dynamicName.GetRange();
            Console.WriteLine($"Dynamic range refers to: {dynRange.RefersTo}");

            // Ensure output directory exists
            string outputPath = "DynamicNamedRange.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook (lifecycle: save)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}