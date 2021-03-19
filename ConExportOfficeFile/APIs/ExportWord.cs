using Aspose.Words;
using Models.ConExportOfficeFile;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace APIs.ConExportOfficeFile
{
    /// <summary>
    /// 匯出 Word
    /// </summary>
    public class ExportWord
    {
        /// <summary>
        /// 匯出
        /// </summary>
        /// <param name="data">資料</param>
        /// <param name="isUseMonth">是否依照月份拆分</param>
        public void Export(List<FileModel> data, bool isUseMonth)
        {
            // https://www.itread01.com/content/1545816163.html

            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            if (isUseMonth)
            {
                var dates = data.Select(x => new DateTime(x.CRT_DATE.Year, x.CRT_DATE.Month, 1)).Distinct().ToList();
                foreach (var date in dates)
                {
                    var fileData = data.Where(x => x.CRT_DATE >= date && x.CRT_DATE < date.AddMonths(1)).ToList();
                    SetDocument(builder, fileData);

                    if (date != dates.Last())
                    {
                        // 換下一頁
                        builder.InsertBreak(BreakType.PageBreak);
                    }
                }
                doc.Save("Files/分頁/WordTest.doc");
            }
            else
            {
                SetDocument(builder, data);
                doc.Save("Files/不分頁/WordTest.doc");
            }
        }

        /// <summary>
        /// 設定文件
        /// </summary>
        /// <param name="builder"></param>
        /// <param name="data">資料</param>
        private void SetDocument(DocumentBuilder builder, List<FileModel> data)
        {
            //builder.Font.NameFarEast = "標楷體"; // 設置字體
            builder.Font.Name = "標楷體"; 

            // 工作頁標題
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center; // 文字至中對齊
            builder.Font.Size = 18; // 字體大小
            builder.Font.Bold = true; // 是否粗體
            builder.Writeln("測試匯出報表"); // 輸入字後換行

            #region 表格(1)
            List<string> titleData = new List<string>() { "列印人員ID", "Leo", "列印人員姓名", "Leo_shi" };
            builder.Font.Size = 12; // 字體大小
            builder.Font.Bold = false; // 是否粗體
            builder.CellFormat.Borders.LineStyle = LineStyle.Single; // 單元格邊框線樣式
            builder.StartTable();
            foreach (var title in titleData)
            {
                SetColumn(builder, ParagraphAlignment.Left, title);
            }
            builder.EndTable();
            #endregion 表格(1)

            builder.Writeln(); // 換行

            #region 表格(2)
            // 設定列頭
            List<string> columnNames = new List<string>() { "帳號", "姓名", "是否刪除", "建立日期" };

            builder.Font.Size = 12; // 字體大小
            builder.Font.Bold = false; // 是否粗體
            builder.CellFormat.Borders.LineStyle = LineStyle.Single; // 單元格邊框線樣式
            builder.StartTable();
            foreach (var columnName in columnNames)
            {
                SetColumn(builder, ParagraphAlignment.Center, columnName);
            }
            builder.EndRow();

            // 設定每行資料
            foreach (var item in data)
            {
                SetColumn(builder, ParagraphAlignment.Left, item.USER_ID);
                SetColumn(builder, ParagraphAlignment.Left, item.USER_NAME);
                SetColumn(builder, ParagraphAlignment.Left, item.DEL_FLG ? "是" : "否");
                SetColumn(builder, ParagraphAlignment.Left, item.CRT_DATE.ToString("yyyy/MM/dd"));

                // 行結束
                builder.EndRow();
            }
            builder.EndTable();
            #endregion 表格(2)
        }

        /// <summary>
        /// 設定表格欄位
        /// </summary>
        /// <param name="builder"></param>
        /// <param name="paragraphAlignment">文字對齊值</param>
        /// <param name="value">值</param>
        private void SetColumn(DocumentBuilder builder, ParagraphAlignment paragraphAlignment, string value)
        {
            // 插入單元格
            builder.InsertCell();
            // 文字對齊值 ( 不能先設定，會影響到上一個 )
            builder.ParagraphFormat.Alignment = paragraphAlignment;
            // 此單元格中填入內容
            builder.Write(value);
        }
    }
}
