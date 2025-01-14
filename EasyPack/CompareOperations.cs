using ClosedXML.Excel;
using Dapper;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using Microsoft.Data.SqlClient;
using System.Diagnostics;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using UtfUnknown;

namespace 打包程式
{
  public class CompareOperations
  {
    /// <summary>
    /// 初始化Excel檔案
    /// </summary>
    /// <param name="ArrayList">分頁列表</param>
    /// <returns></returns>
    public XLWorkbook InitWorkbook(string[] ArrayList)
    {
      XLWorkbook XLWorkbook = new XLWorkbook();
      #region 標頭樣式
      IXLStyle Style = new XLWorkbook().Style;
      Style.Fill.SetBackgroundColor(XLColor.LightGray);
      Style.Font.FontSize = 12;
      Style.Font.FontName = "Arial";
      Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
      Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
      //Style.Alignment.WrapText = true;
      Style.Font.Bold = true;
      Style.Font.FontColor = XLColor.Black;
      Style.Border.SetTopBorder(XLBorderStyleValues.Thin).Border.SetRightBorder(XLBorderStyleValues.Thin)
      .Border.SetBottomBorder(XLBorderStyleValues.Thin).Border.SetLeftBorder(XLBorderStyleValues.Thin);
      #endregion
      foreach (var WorksheetName in ArrayList)
      {
        IXLWorksheet Worksheet = XLWorkbook.AddWorksheet(WorksheetName);

        #region 寫入標頭
        //寫入標頭
        Worksheet.Cell(1, 1).Value = "來源項目";
        Worksheet.Cell(1, 2).Value = "目標項目";
        //套入標頭樣式
        Worksheet.Cell(1, 1).Style = Style;
        Worksheet.Cell(1, 2).Style = Style;
        #endregion

        #region 自適應寬度
        //0.101.0 => Worksheet.Columns().AdjustToContents();     無作用
        //0.102.0 => Worksheet.ColumnsUsed().AdjustToContents(); 無作用        
        Worksheet.ColumnsUsed().AdjustToContents(minWidth: 75, maxWidth: 150);
        Worksheet.RowsUsed().AdjustToContents(minHeight: 20, maxHeight: 30);
        #endregion
      }
      return XLWorkbook;
    }
    /// <summary>
    /// 取得Excel欄位值(字串)
    /// </summary>
    /// <param name="cellValue">Excel欄位值</param>
    /// <returns></returns>
    /// <exception cref="Exception"></exception>
    public static string GetCellValue(XLCellValue cellValue)
    {
      switch (cellValue.Type)
      {
        case XLDataType.Blank:
          return cellValue.GetBlank().ToString();
        case XLDataType.Boolean:
          return cellValue.GetBoolean().ToString();
        case XLDataType.DateTime:
          return cellValue.GetDateTime().ToString("yyyyMMddHHmmss");
        case XLDataType.Error:
          return cellValue.GetError().ToString(); ;
        case XLDataType.Number:
          return cellValue.GetNumber().ToString();
        case XLDataType.Text:
          return cellValue.GetText();
        case XLDataType.TimeSpan:
          return cellValue.GetTimeSpan().ToString();
        default:
          throw new Exception("Cell值轉換失敗");
      }
    }
    /// <summary>
    /// 專案差異比較
    /// </summary>
    /// <param name="excelPath">Excel路徑</param>
    /// <returns></returns>
    public Dictionary<string, Dictionary<string, string>> ProjectComparison(string excelPath)
    {
      //【檔案路徑列表】
      Dictionary<string, Dictionary<string, string>> FileList = new Dictionary<string, Dictionary<string, string>>();
      try
      {
        ConsoleDebug($"比較作業開始");
        //【檔案名稱】
        string ExcelName = Path.GetFileName(excelPath);

        //讀取Excel檔案
        using (XLWorkbook Workbook = new XLWorkbook(excelPath))
        {
          #region 作業表為空值
          if (Workbook == null)
          {
            throw new Exception($"作業表【{ExcelName}】為空值");
          }
          #endregion

          //讀取所有分頁
          IXLWorksheets Worksheets = Workbook.Worksheets;
          #region 查無任何分頁
          if (Worksheets.Count == 0)
          {
            throw new Exception($"作業表【{ExcelName}】內查無分頁");
          }
          #endregion

          foreach (IXLWorksheet worksheet in Worksheets)
          {
            Dictionary<string, string> SheetFiles = new Dictionary<string, string>();
            IXLWorksheet Worksheet = worksheet;
            #region 若本分頁為空，跳至下一個分頁執行
            if (Worksheet == null)
            {
              ConsoleDebug($"作業表【{ExcelName}】找無分頁..跳至下一分頁..");
              continue;
            }
            #endregion

            //獲取資料行集合
            IXLRows Rows = Worksheet.RowsUsed();
            #region 若本分頁沒有任何資料行，跳至下一個分頁執行
            if (Rows == null || Rows.Count() == 0)
            {
              ConsoleDebug($"作業表【{ExcelName}】分頁【{Worksheet.Name}】查無任何資料行..跳至下一分頁..");
              continue;
            }
            #endregion

            ConsoleDebug($"讀取分頁【{Worksheet.Name}】中...");

            //資料為路徑的初始宣告
            string FilePath = string.Empty;
            for (int i = 2; i <= Rows.Count(); i++)
            {
              IXLRow Row = Worksheet.Row(i);
              //獲取該資料行欄位集合
              IXLCells Cells = Row.Cells();
              #region 若該行無任何資料欄位，跳至下一行執行
              if (Cells == null || Cells.Count() == 0)
              {
                ConsoleDebug($"分頁【{Worksheet.Name}】第{i}行查無任何資料欄..跳至下一行..");
                continue;
              }
              #endregion

              #region 來源項目取值
              string SourceCellValue = string.Empty;
              if (Row.Cell(1) != null && Row.Cell(1).Value.IsText)
              {
                SourceCellValue = GetCellValue(Row.Cell(1).Value);
              }
              #endregion

              #region 目標項目取值
              string TargetCellValue = string.Empty;
              if (Row.Cell(2) != null && Row.Cell(2).Value.IsText)
              {
                TargetCellValue = GetCellValue(Row.Cell(2).Value);
              }
              #endregion

              #region 目標項目是否為新值
              string TargetIsNewCellValue = string.Empty;
              if (Row.Cell(3) != null && Row.Cell(3).Value.IsText)
              {
                TargetIsNewCellValue = GetCellValue(Row.Cell(3).Value);
              }
              #endregion

              #region 判斷是否為路徑資料
              bool IsPath = SourceCellValue.StartsWith("$/") || TargetCellValue.StartsWith("$/");
              if (IsPath)
              {
                FilePath = TargetCellValue.Replace("$/", "");
                //跳至下一行
                ConsoleDebug($"第{i}行，確認為路徑:{FilePath}");
                continue;
              }
              #endregion

              #region 【狀況一】來源與目標值皆取不到有效值
              if (String.IsNullOrWhiteSpace(SourceCellValue) && String.IsNullOrEmpty(TargetCellValue))
              {
                //往下一行找
                continue;
              }
              #endregion
              #region 【狀況二】僅有來源值
              else if (!String.IsNullOrWhiteSpace(SourceCellValue) && String.IsNullOrEmpty(TargetCellValue))
              {
                ConsoleDebug($"分頁【{Worksheet.Name}】第{i}行，僅來源值為無效資料");
                continue;
              }
              #endregion
              #region【狀況三】僅有目標值(新檔案)【狀況四】來源與目標值皆有值(更新檔案)
              else
              {
                string TargetPath = $"{FilePath}/{TargetCellValue}";
                SheetFiles.Add(TargetPath, TargetIsNewCellValue);
                ConsoleDebug($"第{i}行確認加入資料:{TargetPath}");
              }
              #endregion
            }
            FileList.Add(Worksheet.Name, SheetFiles);
            ConsoleDebug($"讀取分頁【{Worksheet.Name}】結束...");
          }
          ConsoleDebug($"比較作業結束");
          return FileList;
        }
      }
      catch
      {
        throw;
      }
    }
    /// <summary>
    /// 專案差異檔案複製
    /// </summary>
    /// <param name="SourceFolders">字典<專案,發行區路徑></param>
    /// <param name="TargetFiles">字典<專案,比較過後的差異檔案位置></param>
    /// <returns></returns>
    public void ProjectReplication(Dictionary<string, string> SourceFolders, Dictionary<string, Dictionary<string, string>> TargetFiles)
    {
      #region 來源檔案為空值
      if (TargetFiles == null || TargetFiles.Count == 0)
      {
        throw new Exception("傳入物件無資源項目");
      }
      #endregion
      try
      {
        ConsoleDebug($"檔案作業開始");
        //開始依照傳入的來源去複製檔案
        foreach (var TargetFile in TargetFiles)
        {
          if (TargetFile.Value.Count == 0)
          {
            continue;
          }
          //【專案名稱】
          string ProjectName = TargetFile.Key;
          //【複製VS發行區檔案至打包程式區域】
          string ProjectPath = string.Empty;
          var MatchProject = SourceFolders.Where(n => n.Key == ProjectName);
          if (MatchProject.Any())
          {
            ProjectPath = MatchProject.First().Value;
          }
          else
          {
            throw new Exception($"{ProjectName}來源資料夾無資源項目");
          }

          //【取得當前路徑】
          string ProjectPath_Return = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, $"{ProjectName}【已整理差異發行區】");

          if (Directory.Exists(ProjectPath))
          {
            #region 【先清空再重新建立，確保資料正確性】
            if (Directory.Exists(ProjectPath_Return))
            {
              Directory.Delete(ProjectPath_Return, true);
            }
            Directory.CreateDirectory(ProjectPath_Return);
            ConsoleDebug($"重新建立{ProjectName}【已整理差異發行區】完成");
            #endregion

            #region【取得來源檔案夾中的所有檔案，並複製資料夾以及檔案檔案】
            string[] Files = Directory.GetFiles(ProjectPath, "*.*", SearchOption.AllDirectories);
            ConsoleDebug($"檔案複製中...");
            foreach (string FilePath in Files)
            {
              string RelativeFilePath = FilePath.Substring(ProjectPath.Length + 1);
              string DestinationPath = Path.Combine(ProjectPath_Return, RelativeFilePath);
              Directory.CreateDirectory(Path.GetDirectoryName(DestinationPath));
              File.Copy(FilePath, DestinationPath, true);
            }
            ConsoleDebug($"已將{Path.GetFileName(ProjectPath)}複製於{ProjectName}【已整理差異發行區】");
            #endregion

            #region 【取得需要包含的檔案，並濾除不需要的檔案】
            string[] DifferentFilePath = TargetFile.Value
            .Where(n =>
            {
              string Extension = Path.GetExtension(n.Key);
              if (Extension.Equals(".cs", StringComparison.OrdinalIgnoreCase))
              {
                return false;
              }
              return true;
            })
            .Select(n =>
            {
              List<string> PathItems = n.Key.Split('/').ToList();
              string VersionRootFolderName = (ProjectName != ConfigurationManager.AppSettings["Folder_SQL"] ? ConfigurationManager.AppSettings["VersionRootFolderName_Project"] : ConfigurationManager.AppSettings["VersionRootFolderName_SQL"]);
              int VersionRootIndex = PathItems.IndexOf(VersionRootFolderName);
              int ProjectRootIndex = PathItems.IndexOf(ProjectName, VersionRootIndex + 1);

              if (VersionRootIndex >= 0 && ProjectRootIndex > VersionRootIndex)
              {
                return string.Join(@"\", PathItems.Skip(ProjectRootIndex + 1));
              }
              return string.Empty;
            })
            .Where(path => !string.IsNullOrEmpty(path))
            .ToArray();
            #region 這邊因為為了指令濾除根目錄的bin資料夾，故無整入TraverseDirectories
            //取得根目錄下的資料夾
            string[] SubDirectories = Directory.GetDirectories(ProjectPath_Return);
            //取得根目錄下的檔案
            string[] SubFiles = Directory.GetFiles(ProjectPath_Return);
            foreach (string SubFile in SubFiles)
            {
              string SubFileName = Path.GetFileName(SubFile);
              string PartialPath = SubFile.Substring(ProjectPath_Return.Length + 1);
              if (DifferentFilePath.Any(n => n == PartialPath))
              {
                //包含要留的檔案
                ConsoleDebug($"保留檔案：{Path.GetFileName(SubFileName)}");
                continue;
              }
              //沒包含代表不需要的檔案
              ConsoleDebug($"刪除檔案：{Path.GetFileName(SubFileName)}");
              File.Delete(SubFile);
            }
            foreach (string SubDirectoryPath in SubDirectories)
            {
              string DirectoryName = Path.GetFileName(SubDirectoryPath);
              if (DirectoryName.Equals("bin", StringComparison.OrdinalIgnoreCase))
              {
                //略過bin檔不刪
                continue;
              }
              TraverseDirectories(SubDirectoryPath, DifferentFilePath, ProjectPath_Return);
            }
            #endregion
            #endregion
          }
          else
          {
            throw new Exception($"來源資料夾:【{ProjectName}】不存在");
          }
        }
        ConsoleDebug($"檔案作業結束");
      }
      catch
      {
        throw;
      }
    }
    /// <summary>
    /// 遍巡資料夾及檔案比對保留刪除
    /// </summary>
    /// <param name="CurrentDirectory">目標資料夾路徑</param>
    /// <param name="FilePathList">保留檔案列表</param>
    /// <param name="RemovePathString">根目錄路徑</param>
    public static void TraverseDirectories(string CurrentDirectory, string[] FilePathList, string RemovePathString)
    {
      CompareOperations CompareMgn = new CompareOperations();
      try
      {
        //取得子資料夾集合
        string[] SubDirectories = Directory.GetDirectories(CurrentDirectory);
        //取得檔案集合
        string[] SubFiles = Directory.GetFiles(CurrentDirectory);
        foreach (string SubFile in SubFiles)
        {
          string SubFileName = Path.GetFileName(SubFile);
          string PartialPath = SubFile.Substring(RemovePathString.Length + 1);
          if (FilePathList.Any(n => n == PartialPath))
          {
            //包含要留的檔案
            CompareMgn.ConsoleDebug($"保留檔案：{Path.GetFileName(SubFileName)}");
            continue;
          }
          //沒包含代表不需要的檔案
          CompareMgn.ConsoleDebug($"刪除檔案：{Path.GetFileName(SubFileName)}");
          File.Delete(SubFile);
        }
        //遞迴到沒子資料夾為止
        foreach (string SubdirectoryPath in SubDirectories)
        {
          TraverseDirectories(SubdirectoryPath, FilePathList, RemovePathString);
        }
        if (Directory.GetDirectories(CurrentDirectory).Length == 0 && Directory.GetFiles(CurrentDirectory).Length == 0)
        {
          //刪除空資料夾
          Directory.Delete(CurrentDirectory);
          CompareMgn.ConsoleDebug($"已刪除空資料夾：{Path.GetFileName(CurrentDirectory)}");
        }
      }
      catch
      {
        throw;
      }
    }

    /// <summary>
    /// 處理SQL差異檔案語法分類
    /// </summary>
    /// <returns></returns>
    public dynamic SQLProcess(Dictionary<string, string> FileData, bool SecondaryVerification)
    {
      #region 宣告回傳物件
      dynamic ReturnObj = new ExpandoObject();
      ReturnObj.Status = false;
      ReturnObj.Message = "初始回傳";
      ReturnObj.Data = null;
      #endregion
      #region 宣告初始物件
      string[] Files;
      string Create = ConfigurationManager.AppSettings["Create"];
      string Data = ConfigurationManager.AppSettings["Data"];
      Dictionary<string, string> Name = new Dictionary<string, string>
      {
        {"C", Create},
        {"D", Data}
      };

      #endregion
      #region 正則表達式
      Regex CreateTest = new Regex(@"CREATE\s+TABLE(?!.*[#@]|.*##)|EXEC\s+sys.sp_addextendedproperty|ALTER\s+TABLE(?!.*[#@]|.*##)|EXEC\s+SP_Add_ColumnDesc|CREATE\s+TRIGGER|ALTER\s+TRIGGER", RegexOptions.IgnoreCase);
      Regex DeleteTest = new Regex(@"USE\s+\[WINE95\];|USE\s+\[WINE2016\];|USE\s+\[WINE2022\];|USE\s+\[WINE95\]\s+;|USE\s+\[WINE2016\]\s+;|USE\s+\[WINE2022\]\s+;|USE\s+\[WINE95\]|USE\s+\[WINE2016\]|USE\s+\[WINE2022\]|USE\s+WINE95|USE\s+WINE2016|USE\s+WINE2022", RegexOptions.IgnoreCase);
      Regex IsProcdure = new Regex(@"(CREATE|ALTER) (PROCEDURE|PROC|) ([^\s]+)", RegexOptions.IgnoreCase);
      Regex IsView = new Regex(@"(CREATE|ALTER) (VIEW) ([^\s]+)", RegexOptions.IgnoreCase);
      Regex IsFunction = new Regex(@"(CREATE|ALTER) (FUNCTION) ([^\s]+)", RegexOptions.IgnoreCase);
      Regex IsTrigger = new Regex(@"(CREATE|ALTER) (TRIGGER) ([^\s]+)", RegexOptions.IgnoreCase);
      #endregion
      try
      {
        ConsoleDebug("開始處理SQL分類...");
        //取得SQL差異分類資料夾
        string SQL_Path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, $"{ConfigurationManager.AppSettings["Folder_SQL"]}【已整理差異發行區】");
        string SQL_RunPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, ConfigurationManager.AppSettings["SQL_Run"]);

        #region 判斷是否處理SQL語法處理
        if (!Directory.Exists(SQL_Path))
        {
          ConsoleDebug($"{ConfigurationManager.AppSettings["Folder_SQL"]}【已整理差異發行區】資料夾不存在，無需處理SQL分類");
          ReturnObj.Status = true;
          ReturnObj.Message = "";
          return ReturnObj;
        }
        Files = Directory.GetFiles(SQL_Path, "*.sql", SearchOption.AllDirectories);
        if (Files.Length == 0)
        {
          ConsoleDebug($"{ConfigurationManager.AppSettings["Folder_SQL"]}【已整理差異發行區】資料夾內無檔案，無需處理SQL分類");
          ReturnObj.Status = true;
          ReturnObj.Message = "";
          return ReturnObj;
        }
        if (FileData.Count != Files.Length)
        {
          ReturnObj.Status = false;
          ReturnObj.Message = $"{ConfigurationManager.AppSettings["Folder_SQL"]}【已整理差異發行區】擁有檔案數與實際作業表項目不符";
          return ReturnObj;
        }
        string[] OriFiles = FileData.Keys.Select(n => Path.GetFileName(n)).ToArray();
        if (Files.Any(n => !OriFiles.Contains(Path.GetFileName(n))))
        {
          ReturnObj.Status = false;
          ReturnObj.Message = $"{ConfigurationManager.AppSettings["Folder_SQL"]}【已整理差異發行區】其中有檔案未在實際作業表項目中";
          return ReturnObj;
        }
        #endregion

        #region 重建SQL執行資料夾
        if (Directory.Exists(SQL_RunPath))
        {
          Directory.Delete(SQL_RunPath, true);
          ConsoleDebug($"已刪除{ConfigurationManager.AppSettings["SQL_Run"]}資料夾及其內容");
        }
        Directory.CreateDirectory(SQL_RunPath);
        ConsoleDebug($"{ConfigurationManager.AppSettings["SQL_Run"]}資料夾重建成功");
        #endregion

        foreach (string FilePath in Files)
        {
          #region 姿君用法
          //string Code = GetFileEncoding(FilePath);
          #endregion
          //套件用法
          string Code = CharsetDetector.DetectFromFile(FilePath).Detected.EncodingName;
          #region 開始各別寫入
          using (StreamReader streamReader = new StreamReader(FilePath, Encoding.GetEncoding(Code)))
          {
            string Type = string.Empty;
            string FileName = string.Empty;
            bool IsNew = FileData.Where(n => Path.GetFileName(n.Key) == Path.GetFileName(FilePath)).First().Value != "是";
            string Content = DeleteTest.Replace(streamReader.ReadToEnd(), ""); // 讀取所有內容到字串中

            #region 二次驗證【確認連接字串有效才做驗證】
            if (SecondaryVerification)
            {
              try
              {
                ConsoleDebug($"進行資料庫二次驗證中...");
                string QueryString = string.Empty;
                if (IsProcdure.IsMatch(Content))
                {
                  QueryString = "SELECT * FROM sys.procedures WHERE name = @SQLName";
                }
                else if (IsView.IsMatch(Content))
                {
                  QueryString = "SELECT * FROM sys.views WHERE name = @SQLName";
                }
                else if (IsFunction.IsMatch(Content))
                {
                  QueryString = "SELECT * FROM sys.objects WHERE name = @SQLName AND TYPE IN ('FN', 'FS', 'TF', 'IF', 'FT')";
                }
                else if (IsTrigger.IsMatch(Content))
                {
                  QueryString = "SELECT * FROM sys.triggers WHERE name = @SQLName";
                }
                else
                {
                  throw new Exception("無法判斷語法");
                }
                string SQLName = Path.GetFileName(FilePath).Split('.')[1];
                using (SqlConnection Connection = new SqlConnection(ConfigurationManager.ConnectionStrings["DataBase"].ConnectionString))
                {

                  Connection.Open();
                  if (Connection.State == ConnectionState.Open)
                  {
                    IsNew = Connection.Query(QueryString, new { SQLName = SQLName }).Count() == 0;
                  }
                }
              }
              catch(Exception ex)
              {
                //如果發生錯誤，就不二次驗證了依照Excel的驗證為準
                ConsoleDebug($"資料庫二次驗證發生錯誤...已忽略二次驗證結果。");
                ConsoleDebug($"資料庫二次驗證錯誤訊息：{ex.Message}");
                ConsoleDebug(ex.StackTrace);
              }
            }
            else
            {
              ConsoleDebug($"略過資料庫二次驗證...");
            }
            #endregion

            #region 分辨語法類型並更正錯誤語法
            if (CreateTest.IsMatch(Content))
            {
              Type = "C";
            }
            else
            {
              Type = "D";
            }
            Name.TryGetValue(Type, out FileName);
            if (IsNew)
            {
              //將Alter修正為Create
              Content = Regex.Replace(Content, @"ALTER\s+PROC\s+|ALTER\s+PROCEDURE\s+", "CREATE PROCEDURE ", RegexOptions.IgnoreCase);
              Content = Regex.Replace(Content, @"ALTER\s+FUNCTION", "CREATE FUNCTION", RegexOptions.IgnoreCase);
              Content = Regex.Replace(Content, @"ALTER\s+VIEW", "CREATE VIEW", RegexOptions.IgnoreCase);
            }
            else
            {
              //將Create修正為Alter
              Content = Regex.Replace(Content, @"CREATE\s+PROC\s+|CREATE\s+PROCEDURE\s+", "ALTER PROCEDURE ", RegexOptions.IgnoreCase);
              Content = Regex.Replace(Content, @"CREATE\s+FUNCTION", "ALTER FUNCTION", RegexOptions.IgnoreCase);
              Content = Regex.Replace(Content, @"CREATE\s+VIEW", "ALTER VIEW", RegexOptions.IgnoreCase);
            }
            #endregion

            using (FileStream fileStream = new FileStream(Path.Combine(SQL_RunPath, $"{FileName}.sql"), FileMode.Append, FileAccess.Write))
            {
              using (StreamWriter writer = new StreamWriter(fileStream, Encoding.UTF8))
              {
                writer.WriteLine($"{Content}\nGO");
                ConsoleDebug($"{Path.GetFileName(FilePath)}已成功寫進{FileName}.sql");
              }
            }
          }
          #endregion
        }
        ConsoleDebug("結束處理SQL分類");
        ReturnObj.Status = true;
        ReturnObj.Message = string.Empty;
        return ReturnObj;
      }
      catch
      {
        throw;
      }
    }

    #region 字符集（mapping表）如 Unicode、ASCII [不等於] 字符集的編碼如  utf-8、utf-16、utf-32
    //------------------------------------------
    //UTF編碼規則  |  固定開頭BOM(byte order mark)
    //------------|-----------------------------
    //utf-8_bom	| EF BB BF
    //utf-16_le	| FF FE
    //utf-16_be	| FE FF
    //------------------------------------------
    //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    //*UTF-8判斷規則：
    //*只要是0開頭的byte就表示是ASCII編碼，也就是0xxxxxxx後面的七碼x相容於傳統ASCII編碼，故可以用 Unicode開 ASCII
    //* 非ascii編碼，一律以1開頭再接0，並且最少兩個1
    //*1的數量指出這個字是由幾個byte所組成，如1110xxxxx表示這個字要3個byte
    //* 其後每個子byte都為10開頭
    //*最多每個字4個byte
    //----------------------------------------------------------------------------------------------------------------------------
    //Unicode(16進制) 範圍  |  UTF-8 (2進制)                       |字節(byte) |utf-8判斷規則* 全部byte跑完之後的指標是否等於檔案長度
    //----------------------|------------------------------------------------|----------------------------------------------------
    //0000 0000 ~ 0000 007F | 0xxxxxxx                            |一個       |*只要有非0、非110、非1110、非11110開頭的byte，就是非utf-8
    //0000 0080 ~ 0000 07FF | 110xxxxx 10xxxxxx                   |兩個       |*110開頭的byte，  需有下一個byte，且下一個byte 為10開頭
    //0000 0800 ~ 0000 FFFF | 1110xxxx 10xxxxxx 10xxxxxx          |三個       |*1110開頭的byte， 需有下兩個byte，且下兩個byte皆為10開頭
    //0001 0000 ~ 0010 FFFF | 11110xxx 10xxxxxx 10xxxxxx 10xxxxxx |四個       |*11110開頭的byte，需有下三個byte，且下三個byte皆為10開頭
    //----------------------------------------------------------------------------------------------------------------------------
    //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    //*ANSI並非特定的字符編碼，而是在不同的系統中有不同的系統編碼
    //*在 英文系統 默認ASCII
    //*在 簡體中文 默認GB類
    //*在 繁體中文 默認Big5
    #endregion
    /// <summary>
    /// 取得檔案編碼方式(姿君寫法)
    /// </summary>
    /// <param name="FilePath">檔案來源</param>
    /// <returns></returns>
    public static string GetFileEncoding(string FilePath)
    {
      byte[] bytes = File.ReadAllBytes(FilePath);
      #region 判斷是否為utf-16le、utf-16be、utf-8-bom
      if (bytes[0] == 0xff && bytes[1] == 0xfe)
      {
        return "utf-16le";
      }
      if (bytes[0] == 0xfe && bytes[1] == 0xff)
      {
        return "utf-16be";
      }
      if (bytes[0] == 0xef && bytes[1] == 0xbb && bytes[2] == 0xbf)
      {
        return "utf-8-bom";
      }
      #endregion
      #region 判斷是否為utf-8還是預設的系統編碼
      int index = 0;
      while (index < bytes.Length)
      {
        int one = 0;
        for (int i = 0; i < 8; i++)
        {
          if ((bytes[index] & (0x80 >> i)) == 0)
          {
            break;
          }
          ++one;
        }
        if (one == 0)
        {
          ++index;
          continue;
        }
        if (one == 1)
        {
          return Encoding.Default.BodyName;
        }
        for (int i = 0; i < one - 1; i++)
        {
          ++index;
          if ((bytes[index] & 0x80) != 0x80 || (bytes[index] & (0x80 >> 1)) != 0)
          {
            return Encoding.Default.BodyName;
          }
        }
        ++index;
      }
      return "utf-8";
      #endregion
    }

    /// <summary>
    /// 偵錯方法
    /// </summary>
    /// <param name="Message">訊息文字</param>
    public void ConsoleDebug(string Message = "")
    {
      if (!string.IsNullOrEmpty(Message))
      {
        Console.WriteLine(Message);
        Trace.WriteLine(Message);
        return;
      }
      Console.WriteLine();
    }
  }
}
