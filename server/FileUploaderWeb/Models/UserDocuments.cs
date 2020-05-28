using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace FileUploaderWeb.Models
{
    public class UserDocuments
    {
        public string currentUser { get; set; }
        public List<DocumentsInfo> document { get; set; }
    }
}