using System;
using System.IO;
using System.Net.Http;
using Aspose.Cells;

class DownloadAndLoadExample
{
    static void Main()
    {
        // URL of the sample XLSX file to download
        string fileUrl = "https://example.com/sample.xlsx";

        // Temporary local path for the downloaded file
        string tempPath = Path.Combine(Path.GetTempPath(), "sample.xlsx");

        // Download the XLSX file and save it to the temporary path
        using (HttpClient client = new HttpClient())
        using (HttpResponseMessage response = client.GetAsync(fileUrl).Result)
        using (Stream remoteStream = response.Content.ReadAsStreamAsync().Result)
        using (FileStream fileStream = new FileStream(tempPath, FileMode.Create, FileAccess.Write))
        {
            remoteStream.CopyTo(fileStream);
        }

        // Load the workbook from the downloaded XLSX file
        Workbook workbook = new Workbook(tempPath);

        // Example operation: display the name of the first worksheet
        Console.WriteLine("First worksheet name: " + workbook.Worksheets[0].Name);

        // Save the workbook to a new file using the Save method
        string outputPath = "output.xlsx";
        workbook.Save(outputPath, SaveFormat.Xlsx);

        Console.WriteLine("Workbook saved to: " + outputPath);
    }
}