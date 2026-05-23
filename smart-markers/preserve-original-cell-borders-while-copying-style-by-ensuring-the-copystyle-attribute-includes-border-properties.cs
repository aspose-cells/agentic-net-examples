using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace PreserveBordersDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // ---------- Source cell with custom borders ----------
                // Put some data in source cell A1
                Cell srcCell = cells["A1"];
                srcCell.PutValue("Source");

                // Retrieve the style of the source cell
                Style srcStyle = srcCell.GetStyle();

                // Set border properties on the source style
                srcStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thick;
                srcStyle.Borders[BorderType.TopBorder].Color = Color.Red;

                srcStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thick;
                srcStyle.Borders[BorderType.BottomBorder].Color = Color.Green;

                srcStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thick;
                srcStyle.Borders[BorderType.LeftBorder].Color = Color.Blue;

                srcStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thick;
                srcStyle.Borders[BorderType.RightBorder].Color = Color.Orange;

                // Apply the styled borders to the source cell
                srcCell.SetStyle(srcStyle);

                // ---------- Destination cell ----------
                // Put some data in destination cell B2
                Cell destCell = cells["B2"];
                destCell.PutValue("Destination");

                // Create a new style object for the destination cell
                Style destStyle = workbook.CreateStyle();

                // Copy the entire source style (including borders) to the destination style
                destStyle.Copy(srcStyle); // Copy method copies all style attributes, borders included

                // Apply the copied style to the destination cell
                destCell.SetStyle(destStyle);

                // ---------- Alternative using Range.CopyStyle ----------
                // Define a source range (C1:D3) and apply a different bordered style
                AsposeRange srcRange = cells.CreateRange("C1:D3");
                Style rangeStyle = workbook.CreateStyle();
                rangeStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Medium;
                rangeStyle.Borders[BorderType.TopBorder].Color = Color.Purple;
                rangeStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Medium;
                rangeStyle.Borders[BorderType.BottomBorder].Color = Color.Purple;
                rangeStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Medium;
                rangeStyle.Borders[BorderType.LeftBorder].Color = Color.Purple;
                rangeStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Medium;
                rangeStyle.Borders[BorderType.RightBorder].Color = Color.Purple;
                srcRange.SetStyle(rangeStyle);

                // Destination range where the style will be copied
                AsposeRange destRange = cells.CreateRange("E1:F3");
                // CopyStyle copies all style settings, preserving borders
                destRange.CopyStyle(srcRange);

                // Save the workbook to verify the results
                string outputPath = "PreserveBordersDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}