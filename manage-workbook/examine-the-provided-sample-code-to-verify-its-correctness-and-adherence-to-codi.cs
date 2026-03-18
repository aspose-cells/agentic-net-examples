using System;
using Aspose.Cells;

namespace AsposeCellsVerification
{
    // This program demonstrates the usage of Worksheet.ErrorCheckOptions
    // and validates that the sample code from the documentation works as expected.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: creation)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Retrieve the ErrorCheckOptionCollection from the worksheet
            ErrorCheckOptionCollection options = worksheet.ErrorCheckOptions;

            // Add a new ErrorCheckOption to the collection
            int optionIndex = options.Add();
            ErrorCheckOption option = options[optionIndex];

            // Configure error check settings (sample uses NumberStoredAsText, TextDate, TextNumber)
            option.SetErrorCheck(ErrorCheckType.NumberStoredAsText, false);
            option.SetErrorCheck(ErrorCheckType.TextDate, true);
            option.SetErrorCheck(ErrorCheckType.TextNumber, false);

            // Define a cell area that the error check option will apply to
            // Using the helper method CreateCellArea for readability
            CellArea area = CellArea.CreateCellArea("A1", "B10");
            option.AddRange(area);

            // Verify that the settings have been applied correctly
            Console.WriteLine("NumberStoredAsText check enabled: " + option.IsErrorCheck(ErrorCheckType.NumberStoredAsText));
            Console.WriteLine("TextDate check enabled: " + option.IsErrorCheck(ErrorCheckType.TextDate));
            Console.WriteLine("TextNumber check enabled: " + option.IsErrorCheck(ErrorCheckType.TextNumber));

            // Retrieve and display the number of ranges associated with this option
            int rangeCount = option.GetCountOfRange();
            Console.WriteLine("Number of ranges in ErrorCheckOption: " + rangeCount);

            // Save the workbook (lifecycle: saving)
            // The file name follows the naming convention used in the documentation
            workbook.Save("ErrorCheckVerificationDemo.xlsx");

            // Indicate successful completion
            Console.WriteLine("Workbook saved successfully.");
        }
    }
}