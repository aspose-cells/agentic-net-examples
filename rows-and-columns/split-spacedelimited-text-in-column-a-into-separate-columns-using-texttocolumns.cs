// Title: Use Aspose.Cells TextToColumns in C# to split space‑delimited strings in column A into separate columns
// AI Prompts: Configure TxtLoadOptions with a space separator and apply Cells.TextToColumns to split the data in column A for a defined row range. | Enable TreatConsecutiveDelimitersAsOne to handle multiple spaces when separating column values using Aspose.Cells. | Save the workbook after the TextToColumns operation to persist the newly created columns.
// Common Searches: Aspose.Cells C# split column A space delimited text into multiple columns | How to use TextToColumns with a space separator in Aspose.Cells | C# example for separating space‑separated values in Excel using Aspose.Cells TextToColumns
// Tags: Aspose.Cells TextToColumns space delimiter | C# TxtLoadOptions separator configuration | split column A values into separate Excel columns | process first N rows with TextToColumns | treat consecutive delimiters as one Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a workbook, fills cells A1‑A3 with space‑delimited strings, sets TxtLoadOptions to use a space as the separator (optionally treating consecutive spaces as a single delimiter), calls Cells.TextToColumns on the first three rows of column A, and saves the result to SplitSpaceDelimited.xlsx.
class SplitSpaceDelimited
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate column A with space‑delimited strings
        cells["A1"].PutValue("John Doe 30");
        cells["A2"].PutValue("Jane Smith 28");
        cells["A3"].PutValue("Bob Johnson 45");

        // Configure load options to use space as the delimiter
        TxtLoadOptions loadOptions = new TxtLoadOptions();
        loadOptions.Separator = ' ';                     // space character
        loadOptions.TreatConsecutiveDelimitersAsOne = true; // optional, handles multiple spaces

        // Split the content of column A (row 0, column 0) for the first 3 rows
        int totalRows = 3;
        cells.TextToColumns(0, 0, totalRows, loadOptions);

        // Save the resulting workbook
        workbook.Save("SplitSpaceDelimited.xlsx");
    }
}
