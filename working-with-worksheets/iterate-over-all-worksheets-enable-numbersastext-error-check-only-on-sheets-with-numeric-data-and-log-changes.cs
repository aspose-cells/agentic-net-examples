// Title: Enable “Numbers Stored as Text” Error Check on Worksheets with Numeric Data using Aspose.Cells for .NET (C#)
// Description: Creates or loads a workbook, adds mixed numeric/text data, scans every worksheet for numeric values, activates the NumbersStoredAsText error check via ErrorCheckOptionCollection only on sheets that contain numbers, logs each change, and saves the file.
// Keywords: Aspose.Cells | .NET | C# | NumbersStoredAsText | ErrorCheckOptionCollection | conditional error checking | worksheet numeric detection | log worksheet changes | save workbook | GitHub example
// Common Searches: Aspose.Cells enable Numbers stored as text error check C# | apply error check only on sheets with numbers Aspose.Cells | iterate worksheets and set error check options .NET | conditional error checking in Aspose.Cells workbook | log worksheets where error check was applied
// Developer Intent: Detect numeric cells in each worksheet and turn on the NumbersStoredAsText error check only for those sheets, while recording which worksheets were modified.
// Use Cases: Automatically flag numeric entries stored as text on data‑rich sheets, leaving text‑only sheets untouched. | Build a cleanup routine that applies specific error‑checking rules only where numbers exist, improving performance. | Generate an audit log of worksheets where the NumbersStoredAsText check was enabled for reporting or compliance.
// AI Prompts: Write a C# method that iterates through all worksheets in an Aspose.Cells workbook, detects any numeric cell, and enables the NumbersStoredAsText error check for those sheets only. | Provide code that logs the names of worksheets where the NumbersStoredAsText error check was activated and then saves the workbook to a given path. | Explain how to use ErrorCheckOptionCollection and CellArea in Aspose.Cells to apply a conditional error‑check option to the used range of a worksheet.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates or loads a workbook, adds mixed numeric/text data, scans every worksheet for numeric values, activates the NumbersStoredAsText error check via ErrorCheckOptionCollection only on sheets that contain numbers, logs each change, and saves the file.
    class EnableNumbersAsTextErrorCheck
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (or load an existing one)
                Workbook workbook = new Workbook();

                // -------------------------------------------------
                // Sample data: add numeric and non‑numeric values
                // -------------------------------------------------
                Worksheet ws1 = workbook.Worksheets[0];
                ws1.Name = "DataSheet";
                ws1.Cells["A1"].PutValue(123);          // numeric
                ws1.Cells["A2"].PutValue("Text");       // non‑numeric
                ws1.Cells["B1"].PutValue(45.67);        // numeric

                // Add a new worksheet and obtain its reference
                int ws2Index = workbook.Worksheets.Add();
                Worksheet ws2 = workbook.Worksheets[ws2Index];
                ws2.Name = "TextOnly";
                ws2.Cells["A1"].PutValue("Hello");
                ws2.Cells["A2"].PutValue("World");

                // -------------------------------------------------
                // Iterate over all worksheets
                // -------------------------------------------------
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    bool hasNumeric = false;

                    // Scan cells to detect any numeric value
                    foreach (Cell cell in sheet.Cells)
                    {
                        // Check if the cell's value is a numeric type
                        if (cell.Value is double ||
                            cell.Value is int ||
                            cell.Value is decimal ||
                            cell.Value is float ||
                            cell.Value is long)
                        {
                            hasNumeric = true;
                            break; // No need to continue scanning this sheet
                        }
                    }

                    if (hasNumeric)
                    {
                        // Access the ErrorCheckOptionCollection for the sheet
                        ErrorCheckOptionCollection options = sheet.ErrorCheckOptions;

                        // Add a new ErrorCheckOption
                        int optionIndex = options.Add();
                        ErrorCheckOption option = options[optionIndex];

                        // Enable the "Numbers stored as text" error check
                        option.SetErrorCheck(ErrorCheckType.NumberStoredAsText, true);

                        // Apply the option to the used range of the worksheet
                        int maxRow = sheet.Cells.MaxDataRow;
                        int maxCol = sheet.Cells.MaxDataColumn;
                        option.AddRange(CellArea.CreateCellArea(0, 0, maxRow, maxCol));

                        Console.WriteLine($"Enabled NumbersAsText error check on worksheet '{sheet.Name}'.");
                    }
                    else
                    {
                        Console.WriteLine($"No numeric data found in worksheet '{sheet.Name}'. Skipping error check.");
                    }
                }

                // Save the workbook
                string outputPath = "NumbersAsTextErrorCheckDemo.xlsx";

                try
                {
                    // Ensure the directory exists
                    string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                    if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                    {
                        Directory.CreateDirectory(directory);
                    }

                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
