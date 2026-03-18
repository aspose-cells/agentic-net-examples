using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Index of the worksheet to be copied (e.g., the first sheet)
        int sourceIndex = 0;

        // Create a copy of the source worksheet within the same workbook
        int copiedIndex = workbook.Worksheets.AddCopy(sourceIndex);
        Worksheet copiedSheet = workbook.Worksheets[copiedIndex];

        // Optionally rename the copied worksheet
        copiedSheet.Name = "CopiedSheet";

        // Move the copied worksheet to the desired position (e.g., index 2)
        copiedSheet.MoveTo(2);

        // Save the workbook with the copied and repositioned worksheet
        workbook.Save("output.xlsx");
    }
}