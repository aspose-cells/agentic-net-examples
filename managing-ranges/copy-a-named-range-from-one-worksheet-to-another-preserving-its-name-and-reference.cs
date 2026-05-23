using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

class Program
{
    static void Main()
    {
        try
        {
            // ---------- Create source workbook ----------
            Workbook srcWb = new Workbook();
            Worksheet srcSheet = srcWb.Worksheets[0];
            srcSheet.Name = "Source";

            // Fill some data in the source range A1:B2
            srcSheet.Cells["A1"].PutValue("Item1");
            srcSheet.Cells["B1"].PutValue(10);
            srcSheet.Cells["A2"].PutValue("Item2");
            srcSheet.Cells["B2"].PutValue(20);

            // Define a named range "MyRange" that refers to A1:B2 on the source sheet
            int srcNameIdx = srcWb.Worksheets.Names.Add("MyRange");
            srcWb.Worksheets.Names[srcNameIdx].RefersTo = $"={srcSheet.Name}!$A$1:$B$2";

            // ---------- Create destination workbook ----------
            Workbook destWb = new Workbook();
            Worksheet destSheet = destWb.Worksheets[0];
            destSheet.Name = "Destination";

            // Define source and destination ranges
            AsposeRange srcRange = srcSheet.Cells.CreateRange("A1:B2");
            // Destination range will be placed at C3:D4 (adjust as needed)
            AsposeRange destRange = destSheet.Cells.CreateRange("C3:D4");

            // Copy the data (values, formulas, formatting, etc.) from source to destination
            destRange.Copy(srcRange);

            // ---------- Replicate the named range in the destination workbook ----------
            // Retrieve the original Name object
            Name srcName = srcWb.Worksheets.Names["MyRange"];

            // Build a new RefersTo string that points to the destination range
            string startAddr = CellsHelper.CellIndexToName(destRange.FirstRow, destRange.FirstColumn);
            string endAddr = CellsHelper.CellIndexToName(
                destRange.FirstRow + destRange.RowCount - 1,
                destRange.FirstColumn + destRange.ColumnCount - 1);
            string newRef = $"={destSheet.Name}!{startAddr}:{endAddr}";

            // Add the name to the destination workbook and set its reference
            int destNameIdx = destWb.Worksheets.Names.Add(srcName.Text);
            destWb.Worksheets.Names[destNameIdx].RefersTo = newRef;

            // ---------- Save the workbooks ----------
            string srcPath = "Source.xlsx";
            string destPath = "Destination.xlsx";

            // Ensure we can write the files (overwrite if they exist)
            srcWb.Save(srcPath);
            destWb.Save(destPath);
        }
        catch (Exception ex)
        {
            // Log or display the error details
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}