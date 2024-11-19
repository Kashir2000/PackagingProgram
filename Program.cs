using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;

namespace 打包程式
{
  internal class Program
  {
    //【當前路徑】
    static string BaseDirectory = AppDomain.CurrentDomain.BaseDirectory;
    //【比較資料夾名稱】
    static string FolderName = ConfigurationManager.AppSettings["FolderName"];
    //【比較資料作業表名稱】
    static string ExcelName = ConfigurationManager.AppSettings["ExcelName"];
    //【作業表分頁名稱】
    static string[] Worksheets = ConfigurationManager.AppSettings["Worksheets"].Split(',');
    //【發行區資料夾名稱】
    static Dictionary<string, string> ReleaseArea = new Dictionary<string, string>();

    static void Main(string[] args)
    {
      try
      {
        //【比較資料夾路徑】
        string FolderPath = Path.Combine(BaseDirectory, FolderName);

        #region【檢查目標資料夾是否存在】
        Console.WriteLine("檢查比較項目資料夾是否存在..");
        if (!Directory.Exists(FolderPath))
        {
          //【新增資料夾】
          Directory.CreateDirectory(FolderPath);
          Console.WriteLine("找無比較項目資料夾，已重新建立。");
        }
        #endregion

        //【Excel路徑】
        string ExcelPath = Path.Combine(FolderPath, ExcelName);

        #region【檢查目標Excel是否存在】
        Console.WriteLine($"檢查Excel【{ExcelName}】是否存在..");
        if (!File.Exists(ExcelPath))
        {
          using (XLWorkbook WorkBook = new CompareOperations().InitWorkbook(Worksheets))
          {
            WorkBook.SaveAs(ExcelPath);
          }
          Console.WriteLine($"找無Excel【{ExcelName}】，已重新建立新檔案！請確認資料內容後重新操作一次，請按下任意鍵結束。");
          Console.ReadKey();
          return;
        }
        #endregion

        #region 【開始比較差異檔案】
        dynamic Result = new CompareOperations().ProjectComparison(ExcelPath);
        if (!Result.Status)
        {
          Console.WriteLine($"處理比較差異發生錯誤：{Result.Message}");
          Console.WriteLine("程式中止，請按下任意鍵結束。");
          Console.ReadKey();
          return;
        }
        Dictionary<string, Dictionary<string, string>> SheetFileData = Result.Data;
        //確認取檔路徑
        foreach (var SheetFile in SheetFileData)
        {
          if (SheetFile.Value.Count > 0)
          {
            Console.WriteLine($@"請輸入{SheetFile.Key}專案已發行的有效絕對路徑(未輸入則取固定路徑D:\{SheetFile.Key}發行區)後按下Enter鍵：");
            string CustomPath = Console.ReadLine();
            while (!Directory.Exists(CustomPath))
            {
              if (string.IsNullOrEmpty(CustomPath))
              {
                break;
              }
              Console.WriteLine($@"{CustomPath}不存在，請重新輸入{SheetFile.Key}專案已發行的有效絕對路徑：");
              CustomPath = Console.ReadLine();
            }
            string TagetPath = Directory.Exists(CustomPath) ? CustomPath : $@"D:\{SheetFile.Key}發行區";
            Console.WriteLine($"取值路徑:{TagetPath}");
            ReleaseArea.Add(SheetFile.Key, TagetPath);
          }
        }
        #endregion

        #region 【開始複製差異檔案】
        dynamic Result2 = new CompareOperations().ProjectReplication(ReleaseArea, SheetFileData);
        if (!Result2.Status)
        {
          Console.WriteLine($"處理差異複製中發生錯誤：{Result2.Message}");
          Console.WriteLine("程式中止，請按下任意鍵結束。");
          Console.ReadKey();
          return;
        }
        #endregion

        #region 【處理SQL語法分類】
        Dictionary<string, string> SQL_FileData = new Dictionary<string, string>();
        SheetFileData.TryGetValue(ConfigurationManager.AppSettings["Folder_SQL"], out SQL_FileData);
        Console.Write($@"是否啟用以App.config目標資料庫二次驗證？(Y/N)：");
        string Input = Console.ReadKey().KeyChar.ToString().ToUpper();
        while (Input != "Y" && Input != "N")
        {
          Console.WriteLine();
          Console.Write($@"是否啟用以App.config目標資料庫二次驗證？(Y/N)：");
          Input = Console.ReadKey().KeyChar.ToString().ToUpper();
        }
        bool SecondaryVerification = Input == "Y" ? true : false;
        dynamic Result3 = new CompareOperations().SQLProcess(SQL_FileData, SecondaryVerification);
        if (!Result3.Status)
        {
          Console.WriteLine($"處理SQL語法分類中發生錯誤：{Result3.Message}");
          Console.WriteLine("程式中止，請按下任意鍵結束。");
          Console.ReadKey();
          return;
        }
        #endregion

        Console.WriteLine("處理比較差異完成，請按下任意鍵結束。");
        Console.ReadKey();
      }
      catch (Exception ex)
      {
        Console.WriteLine($"處理比較差異異常：{ex.Message}");
        Console.ReadKey();
        return;
      }
    }
  }
}
