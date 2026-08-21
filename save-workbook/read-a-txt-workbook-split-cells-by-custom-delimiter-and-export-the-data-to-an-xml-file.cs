// Title: C# – Load pipe‑delimited TXT, split columns with TextToColumns, and save as XML using Aspose.Cells
// Description: Shows how to load a TXT file with a custom delimiter via TxtLoadOptions, apply Cells.TextToColumns to separate the data into individual cells, and export the workbook to XML format in .NET.
// Keywords: Aspose.Cells | C# | TxtLoadOptions | custom delimiter | pipe delimited | TextToColumns | export to XML | load txt file | save workbook as XML | txt to xml conversion
// Common Searches: Aspose.Cells load pipe delimited txt | TextToColumns example C# | Save workbook as XML Aspose.Cells | Convert txt to xml using Aspose | How to split custom delimited text in Aspose.Cells
// Developer Intent: Import a delimited TXT file, split its fields into separate cells, and generate an XML workbook.
// Use Cases: Convert legacy pipe‑separated log files into XML for downstream processing. | Transform custom‑delimited exports from a mainframe into XML to feed an XML‑based API. | Automate batch conversion of daily TXT reports to XML within a scheduled .NET service.
// AI Prompts: Generate C# code with Aspose.Cells that loads a semicolon‑delimited TXT, splits the data into columns, and saves the workbook as XML. | Explain the interaction between TxtLoadOptions.Separator and Cells.TextToColumns for parsing custom delimited text in Aspose.Cells. | Provide a step‑by‑step tutorial for converting a tab‑delimited text file to XML using Aspose.Cells for .NET, including handling of empty rows.

using System;
using Aspose.Cells;

namespace AsposeCellsTxtToXmlDemo
{
    // Shows how to load a TXT file with a custom delimiter via TxtLoadOptions, apply Cells.TextToColumns to separate the data into individual cells, and export the workbook to XML format in .NET.
    class Program
    {
        static void Main()
        {
            // Path to the source TXT file (each line contains values separated by a custom delimiter)
            string txtPath = "input.txt";

            // Configure load options with a custom delimiter (e.g., pipe character)
            TxtLoadOptions loadOptions = new TxtLoadOptions();
            loadOptions.Separator = '|';               // you can also use SeparatorString = "|" if preferred

            // Load the TXT file into a workbook using the specified options
            Workbook workbook = new Workbook(txtPath, loadOptions);

            // Access the first worksheet where the data was loaded
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Determine the range to apply TextToColumns:
            // - start at the first row (0)
            // - start at the first column (0) where the combined text resides
            // - totalRows is the number of rows that contain data
            int startRow = 0;
            int startColumn = 0;
            int totalRows = cells.MaxDataRow + 1; // MaxDataRow is zero‑based

            // Apply TextToColumns to split the combined text into separate columns
            // Reuse the same load options (they contain the same separator)
            cells.TextToColumns(startRow, startColumn, totalRows, loadOptions);

            // Export the processed workbook to an XML file.
            // Saving with an .xml extension automatically uses the XML format.
            string xmlOutputPath = "output.xml";
            workbook.Save(xmlOutputPath);

            Console.WriteLine($"TXT file '{txtPath}' has been split and exported to XML at '{xmlOutputPath}'.");
        }
    }
}
