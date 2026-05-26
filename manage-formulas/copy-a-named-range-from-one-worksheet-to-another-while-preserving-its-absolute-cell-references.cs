using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsNamedRangeCopy
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook with two worksheets
                Workbook workbook = new Workbook();
                Worksheet srcSheet = workbook.Worksheets[0];
                srcSheet.Name = "Source";

                Worksheet destSheet = workbook.Worksheets[workbook.Worksheets.Add()];
                destSheet.Name = "Destination";

                // Populate source range A1:B2 with sample data
                srcSheet.Cells["A1"].PutValue("Item");
                srcSheet.Cells["B1"].PutValue(10);
                srcSheet.Cells["A2"].PutValue("Qty");
                srcSheet.Cells["B2"].PutValue(20);

                // Define a named range "MyRange" that refers to the absolute range $A$1:$B$2 on the source sheet
                int nameIndex = workbook.Worksheets.Names.Add("MyRange");
                Name srcName = workbook.Worksheets.Names[nameIndex];
                srcName.RefersTo = $"={srcSheet.Name}!$A$1:$B$2";

                // Retrieve the Range object represented by the named range
                AsposeRange srcRange = srcName.GetRange();

                // Create a destination range on the target worksheet with the same size and address
                AsposeRange destRange = destSheet.Cells.CreateRange(
                    srcRange.FirstRow,
                    srcRange.FirstColumn,
                    srcRange.RowCount,
                    srcRange.ColumnCount);

                // Copy the cell data (including formulas) from source to destination
                destRange.CopyData(srcRange);

                // Create a new named range on the destination sheet that points to the copied range
                int destNameIndex = workbook.Worksheets.Names.Add("MyRangeCopy");
                Name destName = workbook.Worksheets.Names[destNameIndex];

                // Extract the address part (e.g., $A$1:$B$2) from the original RefersTo string
                string addressPart = srcName.RefersTo.Substring(srcName.RefersTo.IndexOf('!') + 1);
                // Build the new RefersTo string with the destination sheet name
                destName.RefersTo = $"={destSheet.Name}!{addressPart}";

                // Ensure the output directory exists
                string outputPath = "NamedRangeCopyResult.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}