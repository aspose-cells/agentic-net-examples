using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsTests
{
    public class FreezePanesTests
    {
        public static void Main()
        {
            try
            {
                new FreezePanesTests().RunFreezePanesTest();
                Console.WriteLine("Test completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Test failed: {ex.Message}");
            }
        }

        public void RunFreezePanesTest()
        {
            // Arrange: create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Act: freeze panes at row index 5 (6th row), column index 0,
            // with 5 frozen rows and 0 frozen columns
            int freezeRow = 5;
            int freezeColumn = 0;
            int frozenRows = 5;
            int frozenColumns = 0;
            sheet.FreezePanes(freezeRow, freezeColumn, frozenRows, frozenColumns);

            // Assert: retrieve freeze information and verify it matches the expected values
            int outRow, outColumn, outFrozenRows, outFrozenColumns;
            bool hasFreeze = sheet.GetFreezedPanes(out outRow, out outColumn, out outFrozenRows, out outFrozenColumns);

            if (!hasFreeze)
                throw new InvalidOperationException("Worksheet should report having frozen panes.");

            if (freezeRow != outRow)
                throw new InvalidOperationException($"Freeze position row mismatch. Expected {freezeRow}, got {outRow}.");

            if (freezeColumn != outColumn)
                throw new InvalidOperationException($"Freeze position column mismatch. Expected {freezeColumn}, got {outColumn}.");

            if (frozenRows != outFrozenRows)
                throw new InvalidOperationException($"Number of frozen rows mismatch. Expected {frozenRows}, got {outFrozenRows}.");

            if (frozenColumns != outFrozenColumns)
                throw new InvalidOperationException($"Number of frozen columns mismatch. Expected {frozenColumns}, got {outFrozenColumns}.");

            // Save the workbook (optional, ensures no runtime errors during save)
            string outputPath = "FreezePanesTestOutput.xlsx";

            try
            {
                workbook.Save(outputPath);
            }
            catch (Exception saveEx)
            {
                throw new IOException($"Failed to save workbook to {outputPath}.", saveEx);
            }
        }
    }
}