// Title: Copy a Named Range Between Worksheets and Preserve Its Name with Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, define a named range on a source sheet, copy the range (including values, formulas, and formatting) to another worksheet, recreate the same named range on the destination sheet, and save the file using Aspose.Cells for C#.
// Keywords: Aspose.Cells copy named range | C# named range duplicate worksheet | preserve named range reference | copy range with formatting Aspose | programmatic named range Aspose.Cells
// Common Searches: copy named range to another sheet Aspose.Cells | preserve named range name when duplicating data | how to copy range with formulas and formatting in .NET | Aspose.Cells create named range programmatically | duplicate table with named range for reporting
// Developer Intent: Duplicate an existing named range on a different worksheet while keeping the original name and updating its reference to the new location.
// Use Cases: Clone a template table to a new sheet while retaining the named range for downstream calculations. | Move a data block to a summary worksheet and keep the named range for dynamic chart sources. | Create a backup of chart source data on another sheet without breaking formulas that rely on the named range.
// AI Prompts: Generate C# code with Aspose.Cells that copies a named range from Sheet1 to Sheet2 and retains the original name. | Show how to copy a named range, including formulas and formatting, and update the RefersTo property for the new sheet using Aspose.Cells. | Explain strategies for handling name conflicts when copying a named range to another worksheet in Aspose.Cells.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, define a named range on a source sheet, copy the range (including values, formulas, and formatting) to another worksheet, recreate the same named range on the destination sheet, and save the file using Aspose.Cells for C#.
    public class CopyNamedRangeDemo
    {
        public static void Run()
        {
            try
            {
                // Create a workbook and a source worksheet
                Workbook workbook = new Workbook();
                Worksheet sourceSheet = workbook.Worksheets[0];
                sourceSheet.Name = "Source";

                // Populate the source range with sample data
                sourceSheet.Cells["A1"].PutValue("Item");
                sourceSheet.Cells["B1"].PutValue("Quantity");
                sourceSheet.Cells["A2"].PutValue("Apple");
                sourceSheet.Cells["B2"].PutValue(10);
                sourceSheet.Cells["A3"].PutValue("Orange");
                sourceSheet.Cells["B3"].PutValue(20);

                // Define a named range "MyRange" that refers to A1:B3 on the source sheet
                int nameIdx = workbook.Worksheets.Names.Add("MyRange");
                Name sourceName = workbook.Worksheets.Names[nameIdx];
                sourceName.RefersTo = $"={sourceSheet.Name}!$A$1:$B$3";

                // Add a destination worksheet
                Worksheet destSheet = workbook.Worksheets.Add("Destination");

                // Retrieve the Range object represented by the named range
                AsposeRange srcRange = sourceName.GetRange();

                // Create a destination range on the destination sheet with the same dimensions
                AsposeRange destRange = destSheet.Cells.CreateRange(
                    srcRange.FirstRow,
                    srcRange.FirstColumn,
                    srcRange.RowCount,
                    srcRange.ColumnCount);

                // Copy data, formulas, formatting, etc., from source range to destination range
                srcRange.Copy(destRange);

                // Create a new named range on the destination sheet that points to the copied range
                int destNameIdx = workbook.Worksheets.Names.Add("MyRange");
                Name destName = workbook.Worksheets.Names[destNameIdx];
                destName.RefersTo = $"={destSheet.Name}!$A$1:$B$3";

                // Save the workbook
                workbook.Save("CopyNamedRangeResult.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            CopyNamedRangeDemo.Run();
        }
    }
}
