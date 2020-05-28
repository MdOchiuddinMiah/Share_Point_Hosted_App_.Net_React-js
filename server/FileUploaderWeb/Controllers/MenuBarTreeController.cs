using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Microsoft.SharePoint.Client;
using FileUploaderWeb.Models;
namespace FileUploaderWeb.Controllers
{
    public class MenuBarTreeController : Controller
    {

        //MenuBarTree/FoldersInfo
        [HttpGet]
        public ActionResult FoldersInfo(string path)
        {
            if (Session["context"] == null)
            {

            }
        
            List<DocumentsInfo> documents = new List<DocumentsInfo>();
            SharePointContext spContext = (SharePointContext)Session["context"];
            ListItemCollection myDiffrentInfos = null;
            using (var clientContext = spContext.CreateUserClientContextForSPHost())
            {
                if (clientContext != null)
                {
                    Web myWeb = clientContext.Web;
                    List differentInfo = myWeb.Lists.GetByTitle("Documents");
                    clientContext.Load(differentInfo);
                    clientContext.Load(differentInfo.RootFolder);
                    CamlQuery camlQuery = new CamlQuery();
                    camlQuery.ViewXml = @"<View>                   
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

                    camlQuery.FolderServerRelativeUrl = path;
                    myDiffrentInfos = differentInfo.GetItems(camlQuery);

                    clientContext.Load(myDiffrentInfos);
                    clientContext.ExecuteQuery();

                    foreach (var item in myDiffrentInfos)
                    {
                        documents.Add(new DocumentsInfo
                        {
                            ID = item["ID"].ToString(),
                            FileLeafRef = item["FileLeafRef"].ToString(),
                            Modified = item["Modified"].ToString(),
                            Created = item["Created"].ToString(),
                            ItemChildCount = Convert.ToInt32(item["ItemChildCount"]),
                            FolderChildCount = Convert.ToInt32(item["FolderChildCount"]),
                            fileType = (item["File_x0020_Type"] == null) ? "" : item["File_x0020_Type"].ToString(),
                            FileRef = item["FileRef"].ToString(),
                            FSObjType = Convert.ToInt32(item["FSObjType"]),
                            ServerRedirectedEmbedUri = (item["ServerRedirectedEmbedUri"] == null) ? "" : item["ServerRedirectedEmbedUri"].ToString()
                        });
                    }

                }
            }



            return Json(documents, JsonRequestBehavior.AllowGet);
        }






    }
}