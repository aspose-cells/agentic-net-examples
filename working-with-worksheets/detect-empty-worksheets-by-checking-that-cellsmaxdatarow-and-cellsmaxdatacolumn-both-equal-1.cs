// Title: Detect empty worksheets in an Excel workbook with Aspose.Cells for .NET using Cells.MaxDataRow and MaxDataColumn
// Common Searches: Aspose.Cells C# check if a worksheet is empty using MaxDataRow | How to find blank sheets in an Excel file with Aspose.Cells .NET | C# code to list worksheets with no data using Cells.MaxDataColumn | Determine empty worksheets in a .xlsx workbook via Aspose.Cells API | Detect worksheets without any cells populated in Aspose.Cells
// Tags: Aspose.Cells empty worksheet detection | Cells.MaxDataRow blank sheet check | Cells.MaxDataColumn empty sheet identification | C# iterate workbook worksheets Aspose.Cells | Excel .xlsx empty sheet detection using Aspose

using System;
using System.IO;
using Aspose.Cells;

namespace DetectEmptyWorksheets
{
    // The example loads a workbook (or creates a sample one), iterates through each worksheet, and uses Cells.MaxDataRow and Cells.MaxDataColumn—both returning -1 for empty sheets—to identify and report empty worksheets in a .xlsx file.
    class Program
    {
        static void Main(string[] args)
        {
            // Determine the input workbook path.
            string workbookPath;
            if (args.Length > 0 && File.Exists(args[0]))
            {
                workbookPath = args[0];
                Console.WriteLine($"Loading workbook from: {workbookPath}");
            }
            else
            {
                // If no valid path is supplied, create a sample workbook for demonstration.
                workbookPath = Path.Combine(Directory.GetCurrentDirectory(), "SampleWorkbook.xlsx");
                Console.WriteLine("No valid input file provided. Creating a sample workbook at:");
                Console.WriteLine(workbookPath);
                CreateSampleWorkbook(workbookPath);
            }

            try
            {
                // Load the workbook.
                Workbook workbook = new Workbook(workbookPath);

                // Iterate through each worksheet and check if it is empty.
                Console.WriteLine("\n--- Empty Worksheet Detection Results ---");
                bool anyEmpty = false;
                for (int i = 0; i < workbook.Worksheets.Count; i++)
                {
                    Worksheet sheet = workbook.Worksheets[i];
                    Cells cells = sheet.Cells;

                    // MaxDataRow and MaxDataColumn return -1 when there is no data.
                    int maxRow = cells.MaxDataRow;
                    int maxCol = cells.MaxDataColumn;

                    if (maxRow == -1 && maxCol == -1)
                    {
                        anyEmpty = true;
                        Console.WriteLine($"Worksheet \"{sheet.Name}\" (Index {i}) is EMPTY.");
                    }
                    else
                    {
                        Console.WriteLine($"Worksheet \"{sheet.Name}\" (Index {i}) contains data. MaxDataRow={maxRow}, MaxDataColumn={maxCol}");
                    }
                }

                if (!anyEmpty)
                {
                    Console.WriteLine("\nNo empty worksheets were found in the workbook.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"\nAn error occurred while processing the workbook: {ex.Message}");
                Console.WriteLine("Stack Trace:");
                Console.WriteLine(ex.StackTrace);
            }
        }

        /// <param name="path">Full file path where the workbook will be saved.</param>
        private static void CreateSampleWorkbook(string path)
        {
            Workbook wb = new Workbook();

            // Worksheet 0 - add some data
            Worksheet ws0 = wb.Worksheets[0];
            ws0.Name = "Sheet1";
            ws0.Cells["A1"].PutValue("Header");
            ws0.Cells["A2"].PutValue(123);

            // Worksheet 1 - leave empty (default empty sheet)
            Worksheet ws1 = wb.Worksheets.Add("Sheet2");

            // Worksheet 2 - add some data
            Worksheet ws2 = wb.Worksheets.Add("Sheet3");
            ws2.Cells["B5"].PutValue(DateTime.Now);

            // Save the workbook
            wb.Save(path, SaveFormat.Xlsx);
        }
    }
}
