using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace FileUploaderWeb.Models
{
    public class DocumentsInfo
    {
        public string ID { get; set; }
        public string FileLeafRef { get; set; }
        public string Modified { get; set; }
        public string Created { get; set; }
        public int ItemChildCount { get; set; }
        public int FolderChildCount { get; set; }
        public string fileType { get; set; }
        public string FileRef { get; set; }
        public int FSObjType { get; set; }
        public string ServerRedirectedEmbedUri { get; set; }
    }
}