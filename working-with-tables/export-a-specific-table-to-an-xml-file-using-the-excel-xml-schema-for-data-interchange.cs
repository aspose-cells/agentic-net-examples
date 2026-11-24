// Title: Export a named Excel table to SpreadsheetML (XML) using Aspose.Cells in C#
// Common Searches: C# Aspose.Cells export a ListObject to an XML SpreadsheetML file | How to save a specific Excel table as XML with Aspose.Cells .NET | Export named table from .xlsx to SpreadsheetML using Aspose.Cells | Aspose.Cells copy table range to a new workbook and save as XML in C# | Save Excel table as XML for data interchange with Aspose.Cells
// Tags: Aspose.Cells ListObject XML export | SpreadsheetML SaveFormat.Xml usage | C# copy Excel table range | export specific worksheet table to XML | Aspose.Cells table to SpreadsheetML conversion

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example loads an existing workbook, locates a named worksheet and ListObject (Excel table), copies the table's values and styles into a fresh workbook, and then saves that workbook as a SpreadsheetML XML file using Aspose.Cells SaveFormat.Xml. It includes file existence checks and basic exception handling.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xml";

            // Verify that the input file exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the source workbook that contains the table.
            Workbook sourceWb = new Workbook(inputPath);

            // Get the worksheet where the table resides (adjust the name as needed).
            Worksheet sourceWs = sourceWb.Worksheets["Sheet1"];
            if (sourceWs == null)
            {
                Console.WriteLine("Worksheet \"Sheet1\" not found in the source workbook.");
                return;
            }

            // Retrieve the table (ListObject) by its name (adjust the name as needed).
            ListObject table = sourceWs.ListObjects["MyTable"];
            if (table == null)
            {
                Console.WriteLine("Table \"MyTable\" not found in the worksheet.");
                return;
            }

            // Create a new workbook that will hold only the exported table.
            Workbook exportWb = new Workbook();
            Worksheet exportWs = exportWb.Worksheets[0];
            exportWs.Name = "ExportedTable";

            // Determine the range of the table (including header) using ListObject properties.
            CellArea tableRange = new CellArea
            {
                StartRow = table.StartRow,
                StartColumn = table.StartColumn,
                EndRow = table.EndRow,
                EndColumn = table.EndColumn
            };

            int rowCount = tableRange.EndRow - tableRange.StartRow + 1;
            int colCount = tableRange.EndColumn - tableRange.StartColumn + 1;

            // Copy the table's data (including styles) to the new worksheet.
            for (int r = 0; r < rowCount; r++)
            {
                for (int c = 0; c < colCount; c++)
                {
                    // Source cell from the original table.
                    Cell srcCell = sourceWs.Cells[tableRange.StartRow + r, tableRange.StartColumn + c];

                    // Destination cell in the export worksheet.
                    Cell destCell = exportWs.Cells[r, c];

                    // Copy value.
                    destCell.PutValue(srcCell.Value);

                    // Copy style.
                    destCell.SetStyle(srcCell.GetStyle());
                }
            }

            // Save the new workbook as an Excel XML file (SpreadsheetML) for data interchange.
            exportWb.Save(outputPath, SaveFormat.Xml);
            Console.WriteLine($"Table exported successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
