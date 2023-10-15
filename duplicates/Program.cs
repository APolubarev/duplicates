

using System;
using OfficeOpenXml;
using System.Text.Json;
using Microsoft.Extensions.Configuration;
using Npgsql;


internal class Program
{
    private static void Main(string[] args)
    {
        Console.WriteLine(DateTime.Now);

        string SettingsfileName = "Settings.json";
        string jsonString = File.ReadAllText(SettingsfileName);
        Settings CurSettings = JsonSerializer.Deserialize<Settings>(jsonString)!;


        String[] f_names = Directory.GetFiles(CurSettings.dir);

        foreach (string f_name in f_names)
        {
            FileInfo f_i = new FileInfo(f_name);
            using (ExcelPackage excelPackage = new ExcelPackage(f_i))
            {

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                ExcelWorksheet firstWorksheet = excelPackage.Workbook.Worksheets[0];

                int Col_article = 0;
                int Col_model = 0;
                int Col_category = 0;

                for (int i = 1; i < firstWorksheet.Cells.Columns; i++)
                {
                    var cur_val = firstWorksheet.Cells[2, i].Value;
                    if (cur_val is null)
                    {
                        break;
                    }

                    string colname = cur_val.ToString();

                    switch (colname)
                    {
                        case "Article INDASTRA characteristics":
                            Col_article = i;
                            break;
                        case "Model":
                            Col_model = i;
                            break;
                        case "Subcategory 1 en":
                            Col_category = i;
                            break;
                        case "Subcategory 2 en":
                            Col_category = i;
                            break;
                        default:
                            break;

                    }

                    // Console.WriteLine("В ячейке: " + colname);   
                }
                // string valB1 = firstWorksheet.Cells[1, 2].Value.ToString();
                // Console.WriteLine("В ячейке: " + valB1);
                using (var conn = new NpgsqlConnection(CurSettings.PGConString))
                {
                    conn.Open();
                    
                    List<string> values = new List<string>(1);
                    for (int n = 3; n < firstWorksheet.Cells.Rows; n++)
                    {

                        string article = GetStrValue(firstWorksheet, n, Col_article);
                        string model = GetStrValue(firstWorksheet, n, Col_model);
                        string category = GetStrValue(firstWorksheet, n, Col_category);
                        string strValue = String.Format("('{0}','{1}','{2}')", article, model, category);
                        values.Add(strValue);
                    }
                    string AllValues = String.Join(",", values);
                    string CommandText = String.Format("INSERT INTO public.duplicates(article_indastra_characteristics, model, category) VALUES {0};", AllValues);
                    using (var command = new NpgsqlCommand(CommandText, conn))
                    {
                        // command.Parameters.AddWithValue("v", AllValues);
                        int nRows = command.ExecuteNonQuery();
                        Console.Out.WriteLine(String.Format("Number of rows updated={0}", nRows));
                    }


                }
            }


        }
        Console.WriteLine(DateTime.Now);
    }
    private static string GetStrValue(OfficeOpenXml.ExcelWorksheet Worksheet, int n, int col)
    {
        var val = Worksheet.Cells[n, col].Value;
        string article;
        if (val is null)
        {
            article = "";
        }
        else
        {
            article = val.ToString();
        }
        return article;
    }
}
public class Settings
{
    public string dir { get; set; }
    public string PGConString { get; set; } = "Server=localhost; User Id=postgres; Database=indastra; Port=5432; Password=1111";

}