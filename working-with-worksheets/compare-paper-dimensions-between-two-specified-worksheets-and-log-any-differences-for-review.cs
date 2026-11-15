// Title: C# program to compare paper size, width, and height of two Excel worksheets with Aspose.Cells and log differences
// Common Searches: how to compare paper size of two worksheets using Aspose.Cells in C# | C# Aspose.Cells example to detect different page setup dimensions between sheets | log mismatched paper width and height of Excel worksheets with Aspose.Cells | compare worksheet page setup properties (PaperSize, PaperWidth, PaperHeight) in .NET | Aspose.Cells code to list differences in sheet printing settings
// Tags: compare worksheet page setup Aspose.Cells | detect paper size mismatch Excel C# | log paper dimension differences Aspose.Cells | retrieve paper width height worksheet Aspose | Aspose.Cells page setup comparison example

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace PaperDimensionComparer
{
    // The sample loads an Excel workbook, checks that two specified worksheets exist, accesses each sheet's PageSetup, and compares the PaperSize enum, PaperWidth, and PaperHeight values. Any mismatches are collected and written to the console for review.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel file (adjust as needed)
            string filePath = @"C:\Path\To\YourWorkbook.xlsx";

            // Names of the worksheets to compare
            string firstSheetName = "Sheet1";
            string secondSheetName = "Sheet2";

            // Verify that the file exists before attempting to load it
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"Error: The file \"{filePath}\" was not found.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(filePath);

                // Retrieve the worksheets; validate their existence
                Worksheet firstSheet = workbook.Worksheets[firstSheetName];
                Worksheet secondSheet = workbook.Worksheets[secondSheetName];

                if (firstSheet == null)
                {
                    Console.WriteLine($"Error: Worksheet \"{firstSheetName}\" does not exist in the workbook.");
                    return;
                }

                if (secondSheet == null)
                {
                    Console.WriteLine($"Error: Worksheet \"{secondSheetName}\" does not exist in the workbook.");
                    return;
                }

                // Get the page setup objects for each sheet
                PageSetup firstSetup = firstSheet.PageSetup;
                PageSetup secondSetup = secondSheet.PageSetup;

                // List to collect any differences found
                List<string> differences = new List<string>();

                // Compare Paper Size (enum)
                if (firstSetup.PaperSize != secondSetup.PaperSize)
                {
                    differences.Add($"PaperSize differs: {firstSheetName} = {firstSetup.PaperSize}, {secondSheetName} = {secondSetup.PaperSize}");
                }

                // Compare Paper Width (points)
                if (Math.Abs(firstSetup.PaperWidth - secondSetup.PaperWidth) > 0.01)
                {
                    differences.Add($"PaperWidth differs: {firstSheetName} = {firstSetup.PaperWidth} pt, {secondSheetName} = {secondSetup.PaperWidth} pt");
                }

                // Compare Paper Height (points)
                if (Math.Abs(firstSetup.PaperHeight - secondSetup.PaperHeight) > 0.01)
                {
                    differences.Add($"PaperHeight differs: {firstSheetName} = {firstSetup.PaperHeight} pt, {secondSheetName} = {secondSetup.PaperHeight} pt");
                }

                // Output the results
                if (differences.Count == 0)
                {
                    Console.WriteLine("No differences in paper dimensions were found between the specified worksheets.");
                }
                else
                {
                    Console.WriteLine("Differences found:");
                    foreach (string diff in differences)
                    {
                        Console.WriteLine(diff);
                    }
                }
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
