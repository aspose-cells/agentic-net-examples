using System;
using System.IO;
using Aspose.Cells;

namespace InsertLogoAfterMergeApp
{
    class InsertLogoAfterMerge
    {
        static void Main()
        {
            try
            {
                // Paths of the workbooks that need to be merged
                string[] sourceFiles = { "File1.xlsx", "File2.xlsx", "File3.xlsx" };

                // Create an empty workbook that will hold the merged result
                Workbook mergedWorkbook = new Workbook();

                // Combine each source workbook into the destination workbook
                foreach (string filePath in sourceFiles)
                {
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"Source file not found: {filePath}. Skipping.");
                        continue;
                    }

                    Workbook source = new Workbook(filePath);
                    mergedWorkbook.Combine(source);
                }

                // Insert the company logo on the first worksheet of the merged workbook
                Worksheet firstSheet = mergedWorkbook.Worksheets[0];
                string logoPath = "CompanyLogo.png";

                if (File.Exists(logoPath))
                {
                    // Add the picture at the top‑left corner (row 0, column 0)
                    firstSheet.Pictures.Add(0, 0, logoPath);
                }
                else
                {
                    Console.WriteLine($"Logo file not found: {logoPath}. Skipping logo insertion.");
                }

                // Save the final workbook with the inserted logo
                mergedWorkbook.Save("MergedWithLogo.xlsx", SaveFormat.Xlsx);
                Console.WriteLine("Merged workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}