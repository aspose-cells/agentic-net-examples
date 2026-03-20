using System;
using Aspose.Cells;

namespace AsposeCellsGridlineToggle
{
    /// <summary>
    /// Demonstrates how to dynamically show or hide gridlines in a worksheet.
    /// </summary>
    public static class GridlineToggleDemo
    {
        /// <summary>
        /// Toggles the gridline visibility of the first worksheet in the specified workbook.
        /// </summary>
        /// <param name="inputPath">Path to the existing workbook file. If the file does not exist, a new workbook will be created.</param>
        /// <param name="outputPath">Path where the modified workbook will be saved.</param>
        /// <param name="showGridlines">True to show gridlines, false to hide them.</param>
        public static void ToggleGridlines(string inputPath, string outputPath, bool showGridlines)
        {
            Workbook workbook;

            // Load existing workbook if the file exists; otherwise create a new one.
            if (System.IO.File.Exists(inputPath))
            {
                // Load rule
                workbook = new Workbook(inputPath);
            }
            else
            {
                // Create rule
                workbook = new Workbook();
            }

            // Access the first worksheet.
            Worksheet worksheet = workbook.Worksheets[0];

            // Set gridline visibility using the IsGridlinesVisible property.
            worksheet.IsGridlinesVisible = showGridlines;

            // Optionally add some sample data so the effect can be observed.
            worksheet.Cells["A1"].PutValue("Gridlines are " + (showGridlines ? "visible" : "hidden"));
            worksheet.AutoFitColumns();

            // Save the workbook using the save rule.
            workbook.Save(outputPath);
        }

        // Example usage
        public static void Main()
        {
            string inputFile = "Sample.xlsx";   // Change as needed
            string outputFileShow = "GridlinesShown.xlsx";
            string outputFileHide = "GridlinesHidden.xlsx";

            // Show gridlines
            ToggleGridlines(inputFile, outputFileShow, true);
            Console.WriteLine($"Workbook saved with gridlines visible: {outputFileShow}");

            // Hide gridlines
            ToggleGridlines(inputFile, outputFileHide, false);
            Console.WriteLine($"Workbook saved with gridlines hidden: {outputFileHide}");
        }
    }
}