using System;
using Aspose.Cells;

class LoadSpecificSheets
{
    static void Main()
    {
        // Path to the source Excel file
        string sourcePath = "input.xlsx";

        // Create a custom load filter that loads only the first and third worksheets (indices 0 and 2)
        LoadFilter filter = new CustomLoadFilter(new int[] { 0, 2 });

        // Configure load options to use the custom filter
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.LoadFilter = filter;

        // Load the workbook with the specified load options – only the selected sheets will be loaded
        Workbook workbook = new Workbook(sourcePath, loadOptions);

        // Display the names of the worksheets that were loaded
        Console.WriteLine("Loaded worksheets:");
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Console.WriteLine(sheet.Name);
        }

        // Save the workbook (containing only the selected sheets) to a new file
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }

    // Custom LoadFilter implementation that overrides SheetsInLoadingOrder
    private class CustomLoadFilter : LoadFilter
    {
        private readonly int[] _sheetsOrder;

        public CustomLoadFilter(int[] sheetsOrder) : base(LoadDataFilterOptions.All)
        {
            _sheetsOrder = sheetsOrder;
        }

        // Return the specific sheet indices and order to be loaded
        public override int[] SheetsInLoadingOrder => _sheetsOrder;
    }
}