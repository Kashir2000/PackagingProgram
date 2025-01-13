using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;

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
    static string[] Worksheets = ConfigurationManager.AppSettings["Worksheets"].Replace("Folder_SQL", ConfigurationManager.AppSettings["Folder_SQL"]).Split(',');
    //【發行區資料夾名稱】
    static Dictionary<string, string> ReleaseArea = new Dictionary<string, string>();

    static void Main(string[] args)
    {
      CompareOperations CompareMgn = new CompareOperations();
      try
      {
        //【比較資料夾路徑】
        string FolderPath = Path.Combine(BaseDirectory, FolderName);

        #region【檢查目標資料夾是否存在】
        CompareMgn.ConsoleDebug("檢查比較項目資料夾是否存在..");
        if (!Directory.Exists(FolderPath))
        {
          //【新增資料夾】
          Directory.CreateDirectory(FolderPath);
          CompareMgn.ConsoleDebug("找無比較項目資料夾，已重新建立。");
        }
        #endregion

        //【Excel路徑】
        string ExcelPath = Path.Combine(FolderPath, ExcelName);

        #region【檢查目標Excel是否存在】
        CompareMgn.ConsoleDebug($"檢查Excel【{ExcelName}】是否存在..");
        if (!File.Exists(ExcelPath))
        {
          using (XLWorkbook WorkBook = CompareMgn.InitWorkbook(Worksheets))
          {
            WorkBook.SaveAs(ExcelPath);
          }
          CompareMgn.ConsoleDebug($"找無Excel【{ExcelName}】，已重新建立新檔案！請確認資料內容後重新操作一次，請按下任意鍵結束。");
          Console.ReadKey();
          return;
        }
        #endregion

        #region 【開始比較差異檔案】
        Dictionary<string, Dictionary<string, string>> SheetFileData = CompareMgn.ProjectComparison(ExcelPath);
        //確認取檔路徑
        foreach (var SheetFile in SheetFileData)
        {
          if (SheetFile.Value.Count > 0)
          {
            string CustomPath;
            do
            {
              CompareMgn.ConsoleDebug($@"請輸入{SheetFile.Key}專案已發行的有效絕對路徑後按下Enter鍵：");
              CustomPath = Console.ReadLine();
              if (string.IsNullOrEmpty(CustomPath))
              {
                CompareMgn.ConsoleDebug($@"尚未輸入{SheetFile.Key}專案已發行的有效絕對路徑，請再試一次！");
              }
              if (!Directory.Exists(CustomPath))
              {
                CompareMgn.ConsoleDebug($@"{CustomPath}不存在，請重新輸入{SheetFile.Key}專案已發行的有效絕對路徑：");
              }
            } while (string.IsNullOrEmpty(CustomPath) || !Directory.Exists(CustomPath));
            CompareMgn.ConsoleDebug($"取值路徑:{CustomPath}");
            ReleaseArea.Add(SheetFile.Key, CustomPath);
          }
        }
        #endregion

        #region 【開始複製差異檔案】
        CompareMgn.ProjectReplication(ReleaseArea, SheetFileData);
        #endregion

        #region 【處理SQL語法分類】
        Dictionary<string, string> SQL_FileData = new Dictionary<string, string>();
        SheetFileData.TryGetValue(ConfigurationManager.AppSettings["Folder_SQL"], out SQL_FileData);
        string Input;
        string[] ValidInputs = { "Y", "N" };
        do
        {
          Console.Write($@"是否啟用以App.config目標資料庫二次驗證？(Y/N)：");
          Input = Console.ReadKey().KeyChar.ToString().ToUpper();
          CompareMgn.ConsoleDebug();
        } while (!ValidInputs.Contains(Input));
        bool SecondaryVerification = Input == "Y";
        CompareMgn.SQLProcess(SQL_FileData, SecondaryVerification);
        #endregion

        CompareMgn.ConsoleDebug("處理比較差異完成，請按下任意鍵結束。");
        Console.ReadKey();
      }
      catch (Exception ex)
      {
        CompareMgn.ConsoleDebug($"處理比較差異作業異常：{ex.Message}");
        CompareMgn.ConsoleDebug($"錯誤堆疊：{ex.StackTrace}");
        CompareMgn.ConsoleDebug("程式中止，請按下任意鍵結束。");
        Console.ReadKey();
        return;
      }
    }
  }
}
