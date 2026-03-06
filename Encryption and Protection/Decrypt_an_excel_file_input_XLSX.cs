using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

public class Program
{
    public static void Main(string[] args)
    {
        if (args.Length != 3)
        {
            Console.WriteLine("Usage: <inputPath> <password> <outputPath>");
            return;
        }

        string inputPath = args[0];
        string password = args[1];
        string outputPath = args[2];

        ExcelDecryptor.Run(inputPath, password, outputPath);
    }
}

public static class ExcelDecryptor
{
    // Decrypts an encrypted XLSX file and saves it without a password.
    // inputPath  : full path to the encrypted Excel file
    // password   : password used to protect the file
    // outputPath : full path where the decrypted file will be saved
    public static void Run(string inputPath, string password, string outputPath)
    {
        // Detect the file format and whether it is encrypted
        FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(inputPath);
        Console.WriteLine($"File is encrypted: {formatInfo.IsEncrypted}");

        // Prepare load options with the password (required for encrypted files)
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Auto);
        loadOptions.Password = password;

        // Load the workbook using the load options
        Workbook workbook = new Workbook(inputPath, loadOptions);

        // Remove the password protection from the workbook settings
        workbook.Settings.Password = null;

        // Save the workbook to the specified output path (now unencrypted)
        workbook.Save(outputPath);
        Console.WriteLine($"Decrypted workbook saved to: {outputPath}");
    }
}