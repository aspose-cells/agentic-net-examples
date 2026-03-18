using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

class Program
{
    static void Main()
    {
        string sourcePath = "input.fods";
        string destPath = "output.csv";

        LoadOptions loadOptions = new LoadOptions(LoadFormat.Fods);
        TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.CSV);

        ConversionUtility.Convert(sourcePath, loadOptions, destPath, saveOptions);

        Console.WriteLine("FODS file has been successfully converted to CSV.");
    }
}