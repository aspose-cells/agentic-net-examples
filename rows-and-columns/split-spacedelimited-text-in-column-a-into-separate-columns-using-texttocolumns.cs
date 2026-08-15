// Title: C# – Split Space‑Delimited Text in Column A Using Aspose.Cells TextToColumns
// Description: Creates a workbook, writes space‑separated strings to cells A1‑A3, configures TxtLoadOptions with a space separator (optionally collapsing multiple spaces), applies Cells.TextToColumns to the first three rows of column A, and saves the result as an XLSX file where each token occupies its own column.
// Keywords: Aspose.Cells | C# | TextToColumns | space delimiter | TxtLoadOptions | split column | consecutive delimiters
// Common Searches: Aspose.Cells split space delimited column C# | TextToColumns example with space separator | Treat multiple spaces as one Aspose.Cells | How to separate names and ages in Excel using Aspose.Cells
// Developer Intent: Use Aspose.Cells TextToColumns to divide space‑separated values in column A into individual columns.
// Use Cases: Transform a single column of full names and ages into separate Name and Age columns for reporting. | Prepare data imported from a space‑delimited text file by parsing each field into its own worksheet column. | Handle records with irregular spacing by enabling TreatConsecutiveDelimitersAsOne in TxtLoadOptions.
// AI Prompts: Generate C# code that uses Aspose.Cells TxtLoadOptions to split space‑delimited text in column A and saves the workbook. | Show an example of TextToColumns with a space separator that treats consecutive spaces as a single delimiter. | Explain how to change the row and column range when applying TextToColumns to other worksheet areas.

using System;
using Aspose.Cells;

// Creates a workbook, writes space‑separated strings to cells A1‑A3, configures TxtLoadOptions with a space separator (optionally collapsing multiple spaces), applies Cells.TextToColumns to the first three rows of column A, and saves the result as an XLSX file where each token occupies its own column.
class SplitSpaceDelimited
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate column A with space‑delimited text
        cells["A1"].PutValue("John Doe 30");
        cells["A2"].PutValue("Jane Smith 28");
        cells["A3"].PutValue("Bob Brown 45");

        // Set up load options to use space as the separator
        TxtLoadOptions loadOptions = new TxtLoadOptions();
        loadOptions.Separator = ' ';                     // space character
        loadOptions.TreatConsecutiveDelimitersAsOne = true; // optional: treat multiple spaces as one

        // Split the text in column A (starting at row 0, column 0) for 3 rows
        cells.TextToColumns(0, 0, 3, loadOptions);

        // Save the workbook with the split data
        workbook.Save("SplitSpaceDelimited.xlsx");
    }
}
