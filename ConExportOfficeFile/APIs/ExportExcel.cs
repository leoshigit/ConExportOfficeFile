using Aspose.Cells;
using Models.ConExportOfficeFile;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace APIs.ConExportOfficeFile
{
    /// <summary>
    /// 匯出 Excel
    /// </summary>
    public class ExportExcel
    {
        /// <summary>
        /// 匯出
        /// </summary>
        /// <param name="data">資料</param>
        /// <param name="isUseMonth">是否依照月份拆分</param>
        public void Export(List<FileModel> data, bool isUseMonth)
        {
            // https://newgoodlooking.pixnet.net/blog/post/110342616
            // https://apireference.aspose.com/cells/net/aspose.cells/bordertype
            // https://apireference.aspose.com/cells/net/aspose.cells/cellbordertype

            Workbook excel = new Workbook(); // 建立空白Excel

            if (isUseMonth)
            {
                var dates = data.Select(x => new DateTime(x.CRT_DATE.Year, x.CRT_DATE.Month, 1)).Distinct().ToList();
                foreach(var date in dates)
                {
                    var fileData = data.Where(x => x.CRT_DATE >= date && x.CRT_DATE < date.AddMonths(1)).ToList();

                    // 如果日期不是第一筆，則新增活頁；如果日期是第一筆，則取第一個活頁
                    int index = date != dates.First() ? excel.Worksheets.Add() : 0;
                    Worksheet sheet = excel.Worksheets[index]; // 取得活頁

                    int year = date.Year - 1911;
                    int month = date.Month;
                    string sheetName = $"{year}年{month}月";

                    SetSheet(excel, sheet, fileData, sheetName);
                }
                excel.Save("Files/分頁/ExcelTest.xls");
            }
            else
            {
                Worksheet sheet = excel.Worksheets[0]; // 取得第一個活頁
                SetSheet(excel, sheet, data, "ExportExcel");
                excel.Save("Files/不分頁/ExcelTest.xls");
            }
        }

        /// <summary>
        /// 設定活頁
        /// </summary>
        /// <param name="excel">檔案</param>
        /// <param name="sheet">活頁</param>
        /// <param name="data">資料</param>
        /// <param name="sheetName">活頁名稱</param>
        private void SetSheet(Workbook excel, Worksheet sheet, List<FileModel> data, string sheetName)
        {
            #region 設定樣式區
            Style titleStyle = excel.CreateStyle();
            titleStyle.Font.Size = 12; // 文字大小
            titleStyle.Font.Name = "標楷體"; // 字型
            titleStyle.Font.IsBold = true; // 粗體
            titleStyle.HorizontalAlignment = TextAlignmentType.Center; // 文字居中

            Style columnTitleStyle = excel.CreateStyle();
            columnTitleStyle.Font.Size = 12; // 文字大小
            columnTitleStyle.Font.Name = "標楷體"; // 字型
            columnTitleStyle.HorizontalAlignment = TextAlignmentType.Center; // 文字居中
            columnTitleStyle.SetBorder(BorderType.TopBorder, CellBorderType.Thick, Color.Black);
            columnTitleStyle.SetBorder(BorderType.BottomBorder, CellBorderType.Thick, Color.Black);
            columnTitleStyle.SetBorder(BorderType.RightBorder, CellBorderType.Thick, Color.Black);
            columnTitleStyle.SetBorder(BorderType.LeftBorder, CellBorderType.Thick, Color.Black);

            Style columnStyle = excel.CreateStyle();
            columnStyle.Font.Size = 12; // 文字大小
            columnStyle.Font.Name = "標楷體"; // 字型
            columnStyle.SetBorder(BorderType.TopBorder, CellBorderType.Thick, Color.Black);
            columnStyle.SetBorder(BorderType.BottomBorder, CellBorderType.Thick, Color.Black);
            columnStyle.SetBorder(BorderType.RightBorder, CellBorderType.Thick, Color.Black);
            columnStyle.SetBorder(BorderType.LeftBorder, CellBorderType.Thick, Color.Black);
            #endregion 設定樣式區

            Cells cells = sheet.Cells; // 取得活頁的每個欄位
            sheet.Name = sheetName; // 活頁名稱

            SetColumn(cells, "A1", titleStyle, "測試匯出報表");

            // 合併儲存格 (第一個列位置, 第一個欄位置, 要合併列的總量, 要合併欄的總量)
            cells.Merge(0, 0, 1, 4);

            SetColumn(cells, "A2", columnStyle, "列印人員ID");
            SetColumn(cells, "B2", columnStyle, "Leo");
            SetColumn(cells, "C2", columnStyle, "列印人員姓名");
            SetColumn(cells, "D2", columnStyle, "Leo_shi");

            // 設定列頭
            SetColumn(cells, "A4", columnTitleStyle, "帳號");
            SetColumn(cells, "B4", columnTitleStyle, "姓名");
            SetColumn(cells, "C4", columnTitleStyle, "是否刪除");
            SetColumn(cells, "D4", columnTitleStyle, "建立日期");

            // 設定每行資料
            int row = 5;
            foreach (var item in data)
            {
                SetColumn(cells, "A" + row, columnTitleStyle, item.USER_ID);
                SetColumn(cells, "B" + row, columnTitleStyle, item.USER_NAME);
                SetColumn(cells, "C" + row, columnTitleStyle, item.DEL_FLG ? "是" : "否");
                SetColumn(cells, "D" + row, columnTitleStyle, item.CRT_DATE.ToString("yyyy/MM/dd"));
                row++;
            }

            sheet.AutoFitColumns(); // 自動調整欄寬
            //sheet.AutoFitRows(); // 自動調整列高
        }

        /// <summary>
        /// 設定欄位
        /// </summary>
        /// <param name="cells">活頁內欄位清單</param>
        /// <param name="column">欄位位置</param>
        /// <param name="style">樣式</param>
        /// <param name="value">值</param>
        private void SetColumn(Cells cells, string column, Style style, string value)
        {
            cells[column].PutValue(value);
            cells[column].SetStyle(style);
        }
    }
}
