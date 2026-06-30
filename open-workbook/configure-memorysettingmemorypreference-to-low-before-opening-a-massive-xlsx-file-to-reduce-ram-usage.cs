using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Configure load options to use low memory mode (MemoryPreference)
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Auto);
        loadOptions.MemorySetting = MemorySetting.MemoryPreference; // reduces RAM usage

        // Open the massive XLSX file with the configured options
        using (FileStream fs = new FileStream("massive.xlsx", FileMode.Open, FileAccess.Read))
        {
            Workbook workbook = new Workbook(fs, loadOptions);

            // Example: read the value of cell A1 from the first worksheet
            Console.WriteLine(workbook.Worksheets[0].Cells["A1"].StringValue);
        }

        // Author note: This snippet demonstrates setting MemorySetting before loading a workbook.
    }
}