using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class HideZeroValuesInQSheets
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (default sheet is added automatically)
                Workbook workbook = new Workbook();

                // Set up demonstration sheets
                Worksheet sheetJan = workbook.Worksheets[0];
                sheetJan.Name = "Q_January";
                sheetJan.Cells["A1"].PutValue(0); // zero value to be hidden

                // Add February sheet and put a zero value
                Worksheet sheetFeb = workbook.Worksheets.Add("Q_February");
                sheetFeb.Cells["B2"].PutValue(0);

                // Add Data sheet and put a zero value (zeros remain visible on this sheet)
                Worksheet sheetData = workbook.Worksheets.Add("Data");
                sheetData.Cells["C3"].PutValue(0);

                // Hide zero values on all worksheets whose name starts with "Q"
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    if (sheet.Name.StartsWith("Q", StringComparison.OrdinalIgnoreCase))
                    {
                        sheet.DisplayZeros = false;
                    }
                }

                // Save the workbook
                string outputPath = "Output_HideZeros.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            HideZeroValuesInQSheets.Run();
        }
    }
}