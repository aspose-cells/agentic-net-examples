using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsErrorCheckReport
{
    class Program
    {
        static void Main(string[] args)
        {
            // Input workbook path (change as needed)
            string inputPath = "InputWorkbook.xlsx";

            // Output report file path
            string reportPath = "ErrorCheckReport.txt";

            // Load the workbook (uses the provided load rule)
            Workbook workbook = new Workbook(inputPath);

            // Prepare a StreamWriter for the report
            using (StreamWriter writer = new StreamWriter(reportPath))
            {
                // Iterate through all worksheets in the workbook
                foreach (Worksheet worksheet in workbook.Worksheets)
                {
                    writer.WriteLine($"Worksheet: {worksheet.Name}");
                    ErrorCheckOptionCollection options = worksheet.ErrorCheckOptions;

                    // Iterate through each ErrorCheckOption in the collection
                    for (int optIndex = 0; optIndex < options.Count; optIndex++)
                    {
                        ErrorCheckOption option = options[optIndex];
                        writer.WriteLine($"  Option #{optIndex}");

                        // List all ranges associated with this option
                        int rangeCount = option.GetCountOfRange();
                        for (int r = 0; r < rangeCount; r++)
                        {
                            CellArea area = option.GetRange(r);
                            writer.WriteLine($"    Range {r}: {area.StartRow},{area.StartColumn} - {area.EndRow},{area.EndColumn}");
                        }

                        // Check each possible ErrorCheckType and log those that are enabled
                        foreach (ErrorCheckType checkType in Enum.GetValues(typeof(ErrorCheckType)))
                        {
                            if (option.IsErrorCheck(checkType))
                            {
                                writer.WriteLine($"    Enabled Check: {checkType}");
                            }
                        }
                    }

                    writer.WriteLine(); // Blank line between worksheets
                }
            }

            // Optionally, save the workbook if any modifications were made (uses the provided save rule)
            // workbook.Save("ModifiedWorkbook.xlsx");
        }
    }
}