using System;
using Aspose.Cells;

public class Program
{
    public static void Main()
    {
        WorkbookMergeDemo.Run();
    }
}

public class WorkbookMergeDemo
{
    public static void Run()
    {
        string[] files = new string[] { "File1.xlsx", "File2.xlsx", "File3.xlsx" };

        Workbook destWorkbook = new Workbook(files[0]);

        for (int i = 1; i < files.Length; i++)
        {
            using (Workbook srcWorkbook = new Workbook(files[i]))
            {
                destWorkbook.Combine(srcWorkbook);
            }
        }

        destWorkbook.Save("CombinedOutput.xlsx", SaveFormat.Xlsx);
        destWorkbook.Dispose();
    }
}