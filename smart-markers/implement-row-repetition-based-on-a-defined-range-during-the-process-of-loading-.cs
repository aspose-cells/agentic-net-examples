using System;
using Aspose.Cells;

namespace RowRepetitionExample
{
    class Program
    {
        static void Main()
        {
            // Path to the source XLSX workbook
            string sourcePath = "input.xlsx";

            // Load the workbook
            Workbook workbook = new Workbook(sourcePath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Define the range of rows that need to be repeated (0‑based indices)
            int firstRowToRepeat = 0;   // e.g., first row (header)
            int lastRowToRepeat = 1;    // e.g., second row
            int rowsToRepeatCount = lastRowToRepeat - firstRowToRepeat + 1;

            // Number of times the defined rows should be repeated
            int repeatTimes = 3;

            // Starting point where the first repetition will be inserted
            // Insert after the original block of rows
            int insertPosition = lastRowToRepeat + 1;

            // Loop to insert and copy the rows the required number of times
            for (int i = 0; i < repeatTimes; i++)
            {
                // Insert blank rows to make space for the copy
                cells.InsertRows(insertPosition, rowsToRepeatCount, true);

                // Copy the source rows into the newly inserted area
                // sourceRowIndex = firstRowToRepeat, destinationRowIndex = insertPosition
                cells.CopyRows(cells, firstRowToRepeat, insertPosition, rowsToRepeatCount);

                // Move the insert position forward for the next iteration
                insertPosition += rowsToRepeatCount;
            }

            // Save the modified workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
        }
    }
}