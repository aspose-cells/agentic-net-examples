using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsErrorHandlingDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the encrypted Excel file
            string filePath = "encrypted.xlsx";

            // Password supplied by the user (could be correct or incorrect)
            string password = "test";

            // Detect file format and encryption status
            FileFormatInfo fileInfo = FileFormatUtil.DetectFileFormat(filePath);

            if (fileInfo.IsEncrypted)
            {
                bool isPasswordCorrect = false;

                // Verify the password without loading the workbook
                try
                {
                    using (FileStream stream = File.OpenRead(filePath))
                    {
                        isPasswordCorrect = FileFormatUtil.VerifyPassword(stream, password);
                    }
                }
                catch (CellsException ex) when (ex.Code == ExceptionType.FileCorrupted)
                {
                    Console.WriteLine("The file is corrupted and cannot be processed.");
                    return;
                }

                if (!isPasswordCorrect)
                {
                    Console.WriteLine("Incorrect password supplied for the encrypted file.");
                    return;
                }

                // Password is correct; attempt to load the workbook
                try
                {
                    LoadOptions loadOptions = new LoadOptions { Password = password };
                    Workbook workbook = new Workbook(filePath, loadOptions);
                    Console.WriteLine("Workbook loaded successfully with the correct password.");
                    // Perform further operations on the workbook here...
                }
                catch (CellsException ex) when (ex.Code == ExceptionType.FileCorrupted)
                {
                    Console.WriteLine("The file appears to be corrupted despite a correct password.");
                }
            }
            else
            {
                // File is not encrypted; load normally
                try
                {
                    Workbook workbook = new Workbook(filePath);
                    Console.WriteLine("Workbook loaded successfully (file is not encrypted).");
                    // Perform further operations on the workbook here...
                }
                catch (CellsException ex) when (ex.Code == ExceptionType.FileCorrupted)
                {
                    Console.WriteLine("The file is corrupted and cannot be opened.");
                }
            }
        }
    }
}