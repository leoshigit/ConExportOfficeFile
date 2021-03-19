using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Models.ConExportOfficeFile
{
    public class FileModel
    {
        public string USER_ID { get; set; }
        public string USER_NAME { get { return USER_ID; } }
        public bool DEL_FLG { get; set; }
        public DateTime CRT_DATE { get; set; }
    }
}
