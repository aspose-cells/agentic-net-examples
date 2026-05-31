using System;
using Aspose.Cells;

namespace AsposeCellsNumbersAsTextDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Put a numeric value as text into cell A1 (this would normally raise the NumbersAsText warning)
            worksheet.Cells["A1"].PutValue("123");

            // Simulate the warning comment that Excel would add for NumbersAsText
            int commentIndex = worksheet.Comments.Add("A1");
            worksheet.Comments[commentIndex].Note = "Number stored as text";

            // Verify that the comment exists before disabling the warning
            Console.WriteLine($"Comments count before disabling warning: {worksheet.Comments.Count}");

            // Disable the NumbersAsText (TextNumber) warning for the range that includes A1
            ErrorCheckOptionCollection errorCheckOptions = worksheet.ErrorCheckOptions;
            int optionIndex = errorCheckOptions.Add();                     // Add a new error‑check option
            ErrorCheckOption errorCheckOption = errorCheckOptions[optionIndex];
            errorCheckOption.SetErrorCheck(ErrorCheckType.TextNumber, false); // Disable NumbersAsText warning
            CellArea cellArea = CellArea.CreateCellArea("A1", "A1");       // Define the range to which the option applies
            errorCheckOption.AddRange(cellArea);

            // After disabling the warning, the comment indicating the issue should be removed.
            // Aspose.Cells automatically removes such comments, but we verify and clean up if needed.
            if (worksheet.Comments.Count > 0)
            {
                // Attempt to remove the comment at A1 (if it still exists)
                worksheet.Comments.RemoveAt("A1");
            }

            // Verify that the comment has been removed
            Console.WriteLine($"Comments count after disabling warning: {worksheet.Comments.Count}");

            // Save the workbook (optional, just to complete the lifecycle)
            workbook.Save("NumbersAsTextDisabled.xlsx", SaveFormat.Xlsx);
        }
    }
}