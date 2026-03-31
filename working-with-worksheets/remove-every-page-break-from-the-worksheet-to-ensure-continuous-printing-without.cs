using Aspose.Cells;

class Program
{
    static void Main()
    {
        Run();
    }

    public static void Run()
    {
        // Load an existing workbook (replace the path with your file)
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Remove all horizontal page breaks
        worksheet.HorizontalPageBreaks.Clear();

        // Remove all vertical page breaks
        worksheet.VerticalPageBreaks.Clear();

        // Save the workbook after removing page breaks
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}