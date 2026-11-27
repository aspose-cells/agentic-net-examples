// Title: Insert a company logo into the first worksheet of a merged Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells add company logo to first sheet after merging multiple Excel files | how to place an image at cell A1 in a merged workbook with Aspose.Cells | merge several .xlsx files and embed a logo on the first worksheet using Aspose.Cells | Aspose.Cells Pictures.Add example for inserting a logo into a combined workbook
// Tags: Aspose.Cells insert picture into worksheet | merge Excel workbooks with Aspose.Cells C# | add company logo to first sheet after workbook merge | Pictures.Add method Aspose.Cells example | save merged workbook as XLSX with logo

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace MergeWorkbooksWithLogo
{
    // The C# program validates command‑line arguments, loads the first workbook as the base, copies worksheets from additional workbooks, inserts a logo image at cell A1 of the first worksheet using the Pictures.Add method, and saves the merged workbook as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Validate arguments
                if (args.Length < 3)
                {
                    Console.WriteLine("Insufficient arguments.");
                    Console.WriteLine("Usage: <logoImagePath> <outputWorkbookPath> <inputWorkbook1> [<inputWorkbook2> ...]");
                    return;
                }

                string logoPath = args[0];
                string outputPath = args[1];
                string[] inputWorkbooks = new string[args.Length - 2];
                Array.Copy(args, 2, inputWorkbooks, 0, inputWorkbooks.Length);

                // Verify logo file exists
                if (!File.Exists(logoPath))
                {
                    Console.WriteLine($"Logo image not found: {logoPath}");
                    return;
                }

                // Verify each input workbook exists
                foreach (var wbPath in inputWorkbooks)
                {
                    if (!File.Exists(wbPath))
                    {
                        Console.WriteLine($"Input workbook not found: {wbPath}");
                        return;
                    }
                }

                // Load the first workbook as the base for merging
                Console.WriteLine($"Loading base workbook: {inputWorkbooks[0]}");
                Workbook mergedWorkbook = new Workbook(inputWorkbooks[0]);

                // Merge remaining workbooks by copying their worksheets
                for (int i = 1; i < inputWorkbooks.Length; i++)
                {
                    try
                    {
                        Console.WriteLine($"Merging workbook: {inputWorkbooks[i]}");
                        Workbook wbToMerge = new Workbook(inputWorkbooks[i]);

                        foreach (Worksheet srcSheet in wbToMerge.Worksheets)
                        {
                            // Add a copy of each worksheet to the merged workbook by name
                            mergedWorkbook.Worksheets.AddCopy(srcSheet.Name);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to merge workbook '{inputWorkbooks[i]}': {ex.Message}");
                    }
                }

                // Insert the logo into the first worksheet of the merged workbook
                Worksheet firstSheet = mergedWorkbook.Worksheets[0];
                Console.WriteLine($"Inserting logo into worksheet: {firstSheet.Name}");

                try
                {
                    // Add picture at cell A1 (row 0, column 0) – Add returns the picture index
                    int pictureIndex = firstSheet.Pictures.Add(0, 0, logoPath);
                    Picture logoPicture = firstSheet.Pictures[pictureIndex];

                    // Optional: set picture size
                    // logoPicture.Width = 150;
                    // logoPicture.Height = 50;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to insert logo: {ex.Message}");
                }

                // Save the merged workbook with the logo
                Console.WriteLine($"Saving merged workbook to: {outputPath}");
                mergedWorkbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine("Operation completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred during processing:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
