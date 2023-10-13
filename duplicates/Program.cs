// See https://aka.ms/new-console-template for more information
using System;
using OfficeOpenXml;


internal class Program
{
    private static void Main(string[] args)
    {
        FileInfo fi = new FileInfo(@"c:\Jod\C#\Worm_Gearbox.xlsx");
        using (ExcelPackage excelPackage = new ExcelPackage(fi))
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelWorksheet firstWorksheet = excelPackage.Workbook.Worksheets[1];


            string valB1 = firstWorksheet.Cells[1, 2].Value.ToString();
            Console.WriteLine("В ячейке: " + valB1);
            
        }
    }
}