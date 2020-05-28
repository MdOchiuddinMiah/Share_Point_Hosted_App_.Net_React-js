using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Security;
using FileUploaderWeb.Models;

namespace FileUploaderWeb.Controllers
{
    public class HomeController : Controller
    {
        [SharePointContextFilter]
        public ActionResult Index()
        {     
            var spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);
            Session["context"] = spContext;
            return new FilePathResult("~/index.html", "text/html");
        }

        
        [HttpGet]
        public ActionResult MenuInfo()
        {
            User spUser = null;
            List<DocumentsInfo> documents = new List<DocumentsInfo>();
            UserDocuments info = new UserDocuments();

            SharePointContext spContext=null;
            if (Session["context"] != null)
            {
              spContext = (SharePointContext)Session["context"];
            }



            using (var clientContext = spContext.CreateUserClientContextForSPHost())
            {
                if (clientContext != null)
                {
                    spUser = clientContext.Web.CurrentUser;
                    clientContext.Load(spUser, user => user.Title);
                    clientContext.ExecuteQuery();
                }
            }


            using (var clientContext = spContext.CreateUserClientContextForSPHost())
            {
                if (clientContext != null)
                {
                    Web myWeb = clientContext.Web;
                    List differentInfo = myWeb.Lists.GetByTitle("Documents");
                    CamlQuery myQuery = new CamlQuery();


                    myQuery.ViewXml = @"<View>                   
                       <Query>
                      <OrderBy>
                      <FieldRef Name='FolderChildCount' Ascending='False' />
                      </OrderBy>
                      </Query>
                    <ViewFields>
                       <FieldRef Name='FileLeafRef' />
                       <FieldRef Name='ID' />
                       <FieldRef Name='Modified' />
                       <FieldRef Name='Created' />
                       <FieldRef Name='ItemChildCount' />
                       <FieldRef Name='FolderChildCount' />
                       <FieldRef Name='File_x0020_Type' />
                       <FieldRef Name='FSObjType' />
                    </ViewFields>
                    <QueryOptions />
                    </View>";

                   
                    ListItemCollection myDiffrentInfos = differentInfo.GetItems(myQuery);
                    clientContext.Load(myDiffrentInfos);
                    clientContext.ExecuteQuery();

                    foreach (var item in myDiffrentInfos)
                    {
                        documents.Add(new DocumentsInfo
                        {
                            ID = item["ID"].ToString(),
                            FileLeafRef = item["FileLeafRef"].ToString(),
                            Modified = item["Modified"].ToString(),
                            Created= item["Created"].ToString(),
                            ItemChildCount = Convert.ToInt32(item["ItemChildCount"]),
                            FolderChildCount = Convert.ToInt32(item["FolderChildCount"]),
                            fileType = (item["File_x0020_Type"] == null) ? "" : item["File_x0020_Type"].ToString(),
                            FileRef = item["FileRef"].ToString(),
                            FSObjType=Convert.ToInt32(item["FSObjType"]),
                            ServerRedirectedEmbedUri = (item["ServerRedirectedEmbedUri"] == null) ? "" : item["ServerRedirectedEmbedUri"].ToString()
                        });
                    }

                }
            }

            info.currentUser = spUser.Title;
            info.document = documents;


            return Json(info, JsonRequestBehavior.AllowGet);
        }


    }
}
