// Title: Convert a pipe‑delimited TXT file to XML by splitting the first column with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads a pipe‑delimited TXT file into an Aspose.Cells Workbook using TxtLoadOptions, splits column A into separate columns with TextToColumns, and saves the workbook as an XML file. | Generate a C# example showing how to set a custom separator in TxtLoadOptions, apply TextToColumns on the first column of a loaded TXT workbook, and export the processed workbook to XML with Aspose.Cells.
// Common Searches: aspnet c# how to use TextToColumns on a txt workbook with Aspose.Cells | convert pipe separated txt to xml using Aspose.Cells .NET | load txt file with custom delimiter and split column A in Aspose.Cells | save workbook as xml after processing txt data in C# Aspose.Cells | Aspose.Cells TxtLoadOptions separator example c#
// Tags: TxtLoadOptions custom separator Aspose.Cells | TextToColumns split column C# Aspose.Cells | export workbook to XML Aspose.Cells | pipe delimited txt to XML conversion Aspose.Cells | process txt workbook Aspose.Cells .NET

using System;
using Aspose.Cells;

namespace AsposeCellsTxtToXmlDemo
{
    // The sample loads a pipe‑delimited TXT file into an Aspose.Cells Workbook using TxtLoadOptions, splits the data in the first column into separate columns via TextToColumns, and then saves the workbook as an XML file.
    class Program
    {
        static void Main()
        {
            // Paths for the input TXT file and the output XML file
            string txtPath = "input.txt";
            string xmlPath = "output.xml";

            // -----------------------------------------------------------------
            // 1. Load the TXT file into a Workbook.
            //    Use TxtLoadOptions to specify the same delimiter that the file uses.
            // -----------------------------------------------------------------
            TxtLoadOptions loadOptions = new TxtLoadOptions();
            // Example custom delimiter – change as needed (e.g., '|', ';', etc.)
            loadOptions.Separator = '|';
            Workbook workbook = new Workbook(txtPath, loadOptions);

            // -----------------------------------------------------------------
            // 2. Split the data in the first column into multiple columns.
            //    The TextToColumns method works on a specific column range.
            // -----------------------------------------------------------------
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Determine how many rows contain data in the first column.
            int totalRows = cells.MaxDataRow + 1; // MaxDataRow is zero‑based.

            // Options for the split operation – set the same delimiter.
            TxtLoadOptions splitOptions = new TxtLoadOptions();
            splitOptions.SeparatorString = "|";

            // Split starting at row 0, column 0 (cell A1).
            cells.TextToColumns(0, 0, totalRows, splitOptions);

            // -----------------------------------------------------------------
            // 3. Export the processed workbook to an XML file.
            //    Saving with an .xml extension automatically creates XML output.
            // -----------------------------------------------------------------
            workbook.Save(xmlPath);

            Console.WriteLine($"TXT file '{txtPath}' has been processed and exported to XML file '{xmlPath}'.");
        }
    }
}
