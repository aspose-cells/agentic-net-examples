// Title: Encrypt an Excel workbook with a password and benchmark opening time using Aspose.Cells for .NET in headless automation
// Common Searches: C# Aspose.Cells encrypt Excel file with password and measure open duration | How to benchmark loading of a password-protected XLSX using Aspose.Cells in a headless .NET app | Measure elapsed time for opening an encrypted workbook with Aspose.Cells LoadOptions.Password | Performance test for opening password protected Excel workbook in automated C# script | Headless automation load encrypted XLSX with Aspose.Cells and get load time
// Tags: encrypt workbook Aspose.Cells | load encrypted XLSX Aspose.Cells | benchmark workbook open time .NET | headless automation Excel loading | password protected Excel performance

using System;
using System.Diagnostics;
using Aspose.Cells;

// The example creates a new workbook, adds data, sets a password via Workbook.Settings.Password, saves it as an XLSX, then reloads the file with LoadOptions.Password while timing the operation using Stopwatch, and finally outputs the elapsed milliseconds.
class Program
{
    static void Main()
    {
        // Create a new workbook (using provided create rule)
        var workbook = new Workbook(); // {CreateWorkbook}

        // Add sample data
        workbook.Worksheets[0].Cells["A1"].PutValue("Sample Data");

        // Encrypt the workbook with a password
        workbook.Settings.Password = "SecretPwd123";

        // Save the encrypted workbook (using provided save rule)
        string filePath = "EncryptedWorkbook.xlsx";
        workbook.Save(filePath, SaveFormat.Xlsx); // {SaveWorkbook}

        // Measure time to open the encrypted workbook in a headless environment
        var stopwatch = Stopwatch.StartNew();

        var loadOptions = new LoadOptions()
        {
            Password = "SecretPwd123"
        };

        // Load the encrypted workbook (using provided load rule)
        var loadedWorkbook = new Workbook(filePath, loadOptions); // {LoadWorkbook}

        stopwatch.Stop();

        Console.WriteLine($"Time to open encrypted workbook: {stopwatch.ElapsedMilliseconds} ms");
    }
}
