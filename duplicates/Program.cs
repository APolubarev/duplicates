

using System;
using OfficeOpenXml;
using System.Text.Json;


internal class Program
{
    private static void Main(string[] args)
    {
        Console.WriteLine(DateTime.Now);

            string SettingsfileName = "Settings.json";
            string jsonString = File.ReadAllText(SettingsfileName);
            Settings CurSettings = JsonSerializer.Deserialize<Settings>(jsonString)!;

        // String[] f_names = Directory.GetFiles(@"c:\Jod\Base\Indastra\Файлы загрузки\");
        String[] f_names = Directory.GetFiles(CurSettings.dir);
        //FileInfo fi = new FileInfo(@"c:\Jod\C#\Worm_Gearbox.xlsx");

        foreach (string f_name in f_names)
        {
            FileInfo f_i = new FileInfo(f_name);
            using (ExcelPackage excelPackage = new ExcelPackage(f_i))
            {
                
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                ExcelWorksheet firstWorksheet = excelPackage.Workbook.Worksheets[0];


                string valB1 = firstWorksheet.Cells[1, 2].Value.ToString();
                Console.WriteLine("В ячейке: " + valB1);

            }


        }
        Console.WriteLine(DateTime.Now);
    }
}
public class Settings{
    public string dir { get; set; }
    
}