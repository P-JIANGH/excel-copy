
using System;
using System.IO;
using Microsoft.Office.Interop.Excel;

namespace ExcelTakeout
{
    class ExcelTakeout
    {
        #region 初期化定数の定義
        const int START_ROW = 2;
        const int START_COLUME = 1;

        const string FILE_START_WITH = "";
        const string FILE_END_WITH = "";
        const string SHEET_NAME = "";

        const string NEW_FILE_NAME = "";
        const string NEW_SHEET_NAME = "";

        public static readonly string[] IGNORE = { ".svn", ".git", "xx" };

        const int EOF = -1;
        #endregion

        public static Application app;
        private Workbooks wbs;
        private Workbook wb;
        private Sheets wss;
        private Worksheet ws;

        private int corsur = START_ROW;

        public ExcelTakeout(string Basepath)
        {
            wbs = app.Workbooks;
            string newFilePath = Basepath + "\\" + NEW_FILE_NAME;
            try
            {
                wb = wbs.Open(newFilePath);
                wss = wb.Worksheets;
                ws = wss.Item[NEW_SHEET_NAME];
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                wbs.Close();
            }
        }

        public void Close()
        {
            wb.Save();
            wbs.Close();
            app.Quit();
        }

        public void Paste(int maxRow, int maxCol, string auth = "")
        {
            Range rng = ws.Cells.Range[ws.Cells[corsur, START_COLUME], ws.Cells[corsur + maxRow, maxCol]];
            rng.PasteSpecial(XlPasteType.xlPasteAll, XlPasteSpecialOperation.XlPasteSpecialOperationNone);
            ws.Cells[corsur + 1, START_COLUME + 3] = auth;
            wb.Save();
            corsur += maxRow;
        }

        ~ExcelTakeout()
        {
            wbs.Close();
            app.Quit();
        }

        private static string FindFile(string Basepath)
        {
            string scanResult = null;
            Console.WriteLine("===>  " + Basepath);
            DirectoryInfo dicInfo = new DirectoryInfo(Basepath);

            if (isIgnore(dicInfo.Name))
            {
                return scanResult;
            }
            foreach (FileInfo file in dicInfo.GetFiles())
            {
                Console.WriteLine("  + " + file);
                if (file.Name.StartsWith(FILE_START_WITH) && file.Name.EndsWith(FILE_END_WITH))
                {
                    scanResult = file.FullName;
                    return scanResult;
                }
            }
            foreach (DirectoryInfo dic in dicInfo.GetDirectories())
            {
                Console.WriteLine("  * " + dic);
                if (isIgnore(dic.Name))
                {
                    continue;
                }
                scanResult = FindFile(dic.FullName);
                if (scanResult == null)
                {
                    continue;
                }
                else
                {
                    return scanResult;
                }
            }
            return scanResult;
        }

        private static bool isIgnore(string dic)
        {
            foreach (string ig in IGNORE)
            {
                if (ig == dic) return true;
            }
            return false;
        }

        static void Main(string[] args)
        {
            var Basepath = args[0];
            app = new Application();
            var wbs = app.Workbooks;
            ExcelTakeout excelop = new ExcelTakeout(Basepath);
            foreach (DirectoryInfo dic in dicInfo.GetDirectories())
            {
                string fileFullPath = FindFile(dic.FullName);
                if (fileFullPath != null)
                {
                    try
                    {
                        Workbook wb = wbs.Open(fileFullPath);
                        Worksheet sheet = (Worksheet) wb.Worksheets[SHEET_NAME];    // match to the major sheet
                        int colCount = sheet.UsedRange.Columns.Count;       // get the total column count
                        int rowCorsur = START_ROW;
                        for (; sheet.Cells[row, START_COLUME].Value != null ||
                            sheet.Cells[row + 1, START_COLUME].Value != null ||
                            sheet.Cells[row + 2, START_COLUME].Value != null; row++);
                        if (--row == START_ROW)
                        {
                            wb.Close();
                            continue;
                        }
                        Console.WriteLine(dic + " -> " + fileFullPath);
                        Range rng = sheet.Range[sheet.Cells[START_ROW, START_COLUME], sheet.Cells[row, colCount]];
                        rng.Copy(Type.Missing);
                        excelop.Paste(row, colCount, dic.Name);
                        wb.Close();
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e);
                        excelop.Close();
                    }
                }
            }
            excelop.Close();
        }
    }
}
