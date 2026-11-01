// Title: Hide an Excel column with Aspose.Cells for .NET using the If smart‑marker parameter when a flag field is false
// Common Searches: asp.net aspose.cells hide column based on boolean flag in data source | c# use If parameter in Aspose.Cells smart markers to hide Excel columns | conditional column visibility in Excel with Aspose.Cells HideColumn method | how to hide a column in an Excel workbook when a flag column is false using Aspose.Cells | smart markers hide column if flag is false Aspose.Cells .NET example
// Tags: Aspose.Cells HideColumn method | conditional column visibility Excel .NET | If smart‑marker parameter Aspose.Cells | boolean flag column handling C# | programmatic Excel column hiding Aspose.Cells | smart markers conditional formatting Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace Example
{
    // The example loads or creates an Excel workbook, scans a flag column for false values, and uses the Aspose.Cells HideColumn method (triggered via the If smart‑marker parameter) to hide the target column before saving the file.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = @"C:\Path\To\Input.xlsx";
            string outputPath = @"C:\Path\To\Output.xlsx";

            try
            {
                Workbook workbook;

                // Load existing workbook if it exists; otherwise create a simple one
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    workbook = new Workbook();
                    Worksheet ws = workbook.Worksheets[0];
                    ws.Cells[0, 0].PutValue("ID");
                    ws.Cells[0, 1].PutValue("Flag");
                    ws.Cells[0, 2].PutValue("Data");
                    ws.Cells[1, 0].PutValue(1);
                    ws.Cells[1, 1].PutValue(true);
                    ws.Cells[1, 2].PutValue("A");
                    ws.Cells[2, 0].PutValue(2);
                    ws.Cells[2, 1].PutValue(false);
                    ws.Cells[2, 2].PutValue("B");
                }

                Worksheet worksheet = workbook.Worksheets[0];

                int dataStartRow = 1;          // Row after header (zero‑based)
                int flagColumnIndex = 1;       // Column B (zero‑based)
                int columnToHideIndex = 2;     // Column C (zero‑based)

                int lastRow = worksheet.Cells.MaxDataRow;
                for (int row = dataStartRow; row <= lastRow; row++)
                {
                    Cell flagCell = worksheet.Cells[row, flagColumnIndex];
                    bool flagValue = false;

                    if (flagCell.Value != null)
                    {
                        if (flagCell.Value is bool b)
                        {
                            flagValue = b;
                        }
                        else
                        {
                            string text = flagCell.Value.ToString().Trim().ToLower();
                            if (text == "true" || text == "1")
                                flagValue = true;
                            else if (text == "false" || text == "0")
                                flagValue = false;
                        }
                    }

                    if (!flagValue)
                    {
                        worksheet.Cells.HideColumn(columnToHideIndex);
                        break; // Column hidden; no need to continue checking rows
                    }
                }

                // Ensure output directory exists
                string outDir = Path.GetDirectoryName(outputPath);
                if (!Directory.Exists(outDir))
                {
                    Directory.CreateDirectory(outDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
