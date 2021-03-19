using APIs.ConExportOfficeFile;
using Models.ConExportOfficeFile;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConExportOfficeFile
{
    public class ConExportOfficeFile
    {
        public void Save()
        {
            // 假資料
            var data = GetFileModels();

            // 不分頁
            new ExportWord().Export(data, false);
            new ExportPDF().Export(data, false);
            new ExportExcel().Export(data, false);

            // 分頁
            new ExportWord().Export(data, true);
            new ExportPDF().Export(data, true);
            new ExportExcel().Export(data, true);
        }

        /// <summary>
        /// 取得假資料
        /// </summary>
        /// <returns></returns>
        private List<FileModel> GetFileModels()
        {
            DateTime today = DateTime.Today;
            List<FileModel> result = new List<FileModel>
            {
                new FileModel(){ USER_ID = "A", DEL_FLG = false, CRT_DATE = today},
                new FileModel(){ USER_ID = "B", DEL_FLG = true, CRT_DATE = today.AddDays(1)},
                new FileModel(){ USER_ID = "C", DEL_FLG = false, CRT_DATE = today.AddDays(2)},
                new FileModel(){ USER_ID = "D", DEL_FLG = false, CRT_DATE = today.AddMonths(1)},
                new FileModel(){ USER_ID = "E", DEL_FLG = true, CRT_DATE = today.AddMonths(1).AddDays(1)},
                new FileModel(){ USER_ID = "F", DEL_FLG = false, CRT_DATE = today.AddMonths(1).AddDays(2)},
            };
            return result;
        }
    }
}
