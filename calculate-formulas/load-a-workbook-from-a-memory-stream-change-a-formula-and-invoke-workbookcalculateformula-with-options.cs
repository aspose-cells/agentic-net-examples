// Title: Load an Excel workbook from a MemoryStream, change cell A1 formula, and recalculate all formulas with custom CalculationOptions using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# load workbook from MemoryStream and recalculate formulas | How to change a cell formula and enable recursive calculation with Aspose.Cells | Set formula =SUM(B1:B10) in Aspose.Cells and ignore errors during calculation | Recalculate all formulas after modifying a workbook loaded from a byte array using Aspose.Cells | Save modified workbook to file after applying CalculationOptions in Aspose.Cells .NET
// Tags: load workbook from memory stream Aspose.Cells | modify cell formula Aspose.Cells C# | calculationoptions recursive ignoreerror Aspose.Cells | recalculate all formulas Aspose.Cells | save workbook to file Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example reads or creates an XLSX file, loads it into an Aspose.Cells Workbook via a MemoryStream, sets the formula "=SUM(B1:B10)" in cell A1, configures CalculationOptions with Recursive and IgnoreError enabled, recalculates all formulas, and saves the updated workbook to output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists; create a simple workbook if it does not.
            byte[] workbookData;
            if (File.Exists(inputPath))
            {
                workbookData = File.ReadAllBytes(inputPath);
            }
            else
            {
                // Create a new workbook with sample data in B1:B10.
                Workbook tempWb = new Workbook();
                Worksheet tempSheet = tempWb.Worksheets[0];
                for (int i = 1; i <= 10; i++)
                {
                    tempSheet.Cells[$"B{i}"].PutValue(i);
                }
                using (MemoryStream ms = new MemoryStream())
                {
                    tempWb.Save(ms, SaveFormat.Xlsx);
                    workbookData = ms.ToArray();
                }
                // Save the generated input file for future runs.
                File.WriteAllBytes(inputPath, workbookData);
            }

            // Load the workbook from a memory stream.
            using (MemoryStream inputStream = new MemoryStream(workbookData))
            {
                Workbook workbook = new Workbook(inputStream);

                // Change the formula of cell A1 in the first worksheet.
                Worksheet sheet = workbook.Worksheets[0];
                Cell targetCell = sheet.Cells["A1"];
                targetCell.Formula = "=SUM(B1:B10)";

                // Set calculation options.
                CalculationOptions calcOptions = new CalculationOptions
                {
                    Recursive = true,    // Recalculate dependent formulas recursively.
                    IgnoreError = true   // Ignore errors during calculation.
                };

                // Recalculate all formulas using the specified options.
                workbook.CalculateFormula(calcOptions);

                // Save the modified workbook to a file.
                using (MemoryStream outputStream = new MemoryStream())
                {
                    workbook.Save(outputStream, SaveFormat.Xlsx);
                    File.WriteAllBytes(outputPath, outputStream.ToArray());
                }
            }

            Console.WriteLine("Workbook processed and saved successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
