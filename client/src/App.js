import React, { Component } from 'react';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import axios from 'axios';
import { initializeIcons } from '@uifabric/icons';
import './App.css';
import { DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { SearchBox } from 'office-ui-fabric-react/lib/SearchBox';
import { Nav } from 'office-ui-fabric-react/lib/Nav';
import { BarLoader } from 'react-spinners';
import { ShowDocuments } from './Components/ShowDocuments'
import { Test } from './Components/Test'
class App extends Component {
  constructor(props) {
    debugger;
    super(props);
    initializeIcons();

    this.state = {
      urlroot: 'https://localhost:44301/',
      userDocuments:{"currentUser":"Kawsar Hamid","document":[{"ID":"19","FileLeafRef":"ImpotantFiles","Modified":"12/4/2017 11:59:50 AM","Created":"12/4/2017 8:22:08 AM","ItemChildCount":2,"FolderChildCount":1,"fileType":"","FileRef":"/sites/khDev/Shared Documents/ImpotantFiles","FSObjType":1,"ServerRedirectedEmbedUri":""},{"ID":"24","FileLeafRef":"Text c","Modified":"12/10/2017 5:31:58 AM","Created":"12/5/2017 10:08:52 AM","ItemChildCount":1,"FolderChildCount":1,"fileType":"","FileRef":"/sites/khDev/Shared Documents/Text c","FSObjType":1,"ServerRedirectedEmbedUri":""},{"ID":"23","FileLeafRef":"Final Exams","Modified":"12/5/2017 9:35:47 AM","Created":"12/5/2017 9:35:47 AM","ItemChildCount":0,"FolderChildCount":0,"fileType":"","FileRef":"/sites/khDev/Shared Documents/Final Exams","FSObjType":1,"ServerRedirectedEmbedUri":""},{"ID":"1","FileLeafRef":"Document.docx","Modified":"5/22/2017 11:07:34 AM","Created":"6/15/2016 7:41:48 AM","ItemChildCount":0,"FolderChildCount":0,"fileType":"docx","FileRef":"/sites/khDev/Shared Documents/Document.docx","FSObjType":0,"ServerRedirectedEmbedUri":"https://iglobe.sharepoint.com/sites/khDev/_layouts/15/WopiFrame.aspx?sourcedoc={88fbec37-39b3-490f-a922-8dda5ef3a1b9}\u0026action=interactivepreview"},{"ID":"2","FileLeafRef":"sample.txt","Modified":"6/15/2016 7:49:51 AM","Created":"6/15/2016 7:49:51 AM","ItemChildCount":0,"FolderChildCount":0,"fileType":"txt","FileRef":"/sites/khDev/Shared Documents/sample.txt","FSObjType":0,"ServerRedirectedEmbedUri":""},{"ID":"9","FileLeafRef":"capture.jpg","Modified":"5/22/2017 10:47:56 AM","Created":"5/22/2017 10:47:56 AM","ItemChildCount":0,"FolderChildCount":0,"fileType":"jpg","FileRef":"/sites/khDev/Shared Documents/capture.jpg","FSObjType":0,"ServerRedirectedEmbedUri":""},{"ID":"11","FileLeafRef":"Approval Test Doc.docx","Modified":"5/22/2017 11:44:42 AM","Created":"5/22/2017 11:44:13 AM","ItemChildCount":0,"FolderChildCount":0,"fileType":"docx","FileRef":"/sites/khDev/Shared Documents/Approval Test Doc.docx","FSObjType":0,"ServerRedirectedEmbedUri":"https://iglobe.sharepoint.com/sites/khDev/_layouts/15/WopiFrame.aspx?sourcedoc={752d72ea-6cf9-4bbb-ab74-5c3adaa8ac4e}\u0026action=interactivepreview"},{"ID":"29","FileLeafRef":"CRM Integration-Help Document.docx","Modified":"12/10/2017 8:57:15 AM","Created":"12/10/2017 8:57:15 AM","ItemChildCount":0,"FolderChildCount":0,"fileType":"docx","FileRef":"/sites/khDev/Shared Documents/CRM Integration-Help Document.docx","FSObjType":0,"ServerRedirectedEmbedUri":"https://iglobe.sharepoint.com/sites/khDev/_layouts/15/WopiFrame.aspx?sourcedoc={5149c8cd-7b3d-4cba-b875-6a6fc0b4116e}\u0026action=interactivepreview"},{"ID":"30","FileLeafRef":"b - Copy (5).pptx","Modified":"12/10/2017 8:59:01 AM","Created":"12/10/2017 8:59:01 AM","ItemChildCount":0,"FolderChildCount":0,"fileType":"pptx","FileRef":"/sites/khDev/Shared Documents/b - Copy (5).pptx","FSObjType":0,"ServerRedirectedEmbedUri":"https://iglobe.sharepoint.com/sites/khDev/_layouts/15/WopiFrame.aspx?sourcedoc={60c0fe55-dee4-429f-815a-c0cee5c453da}\u0026action=interactivepreview"},{"ID":"31","FileLeafRef":"a - Copy.docx","Modified":"12/10/2017 9:00:38 AM","Created":"12/10/2017 9:00:38 AM","ItemChildCount":0,"FolderChildCount":0,"fileType":"docx","FileRef":"/sites/khDev/Shared Documents/a - Copy.docx","FSObjType":0,"ServerRedirectedEmbedUri":"https://iglobe.sharepoint.com/sites/khDev/_layouts/15/WopiFrame.aspx?sourcedoc={4e6db9d7-6554-4c1b-8794-ceee7ebee2ac}\u0026action=interactivepreview"},{"ID":"32","FileLeafRef":"a - Copy (2).docx","Modified":"12/10/2017 9:00:41 AM","Created":"12/10/2017 9:00:41 AM","ItemChildCount":0,"FolderChildCount":0,"fileType":"docx","FileRef":"/sites/khDev/Shared Documents/a - Copy (2).docx","FSObjType":0,"ServerRedirectedEmbedUri":"https://iglobe.sharepoint.com/sites/khDev/_layouts/15/WopiFrame.aspx?sourcedoc={3fc8973d-3529-4eab-bfec-a3a13cbfc7b8}\u0026action=interactivepreview"},{"ID":"33","FileLeafRef":"a - Copy (4).docx","Modified":"12/10/2017 9:00:44 AM","Created":"12/10/2017 9:00:44 AM","ItemChildCount":0,"FolderChildCount":0,"fileType":"docx","FileRef":"/sites/khDev/Shared Documents/a - Copy (4).docx","FSObjType":0,"ServerRedirectedEmbedUri":"https://iglobe.sharepoint.com/sites/khDev/_layouts/15/WopiFrame.aspx?sourcedoc={7d06ecea-a96a-4fc6-b248-05ef971a11da}\u0026action=interactivepreview"},{"ID":"34","FileLeafRef":"a - Copy (3).docx","Modified":"12/10/2017 9:00:48 AM","Created":"12/10/2017 9:00:48 AM","ItemChildCount":0,"FolderChildCount":0,"fileType":"docx","FileRef":"/sites/khDev/Shared Documents/a - Copy (3).docx","FSObjType":0,"ServerRedirectedEmbedUri":"https://iglobe.sharepoint.com/sites/khDev/_layouts/15/WopiFrame.aspx?sourcedoc={c0b13f5b-f666-4b53-b91f-95bb57f849fc}\u0026action=interactivepreview"},{"ID":"35","FileLeafRef":"a.docx","Modified":"12/10/2017 9:00:54 AM","Created":"12/10/2017 9:00:54 AM","ItemChildCount":0,"FolderChildCount":0,"fileType":"docx","FileRef":"/sites/khDev/Shared Documents/a.docx","FSObjType":0,"ServerRedirectedEmbedUri":"https://iglobe.sharepoint.com/sites/khDev/_layouts/15/WopiFrame.aspx?sourcedoc={d3ff7cec-b9f9-4328-98cb-e18eae8e132f}\u0026action=interactivepreview"},{"ID":"36","FileLeafRef":"a - Copy (3).pptx","Modified":"12/10/2017 9:00:58 AM","Created":"12/10/2017 9:00:58 AM","ItemChildCount":0,"FolderChildCount":0,"fileType":"pptx","FileRef":"/sites/khDev/Shared Documents/a - Copy (3).pptx","FSObjType":0,"ServerRedirectedEmbedUri":"https://iglobe.sharepoint.com/sites/khDev/_layouts/15/WopiFrame.aspx?sourcedoc={26eec7ba-05fc-442c-866d-971c3b91a502}\u0026action=interactivepreview"},{"ID":"37","FileLeafRef":"a - Copy (2).pptx","Modified":"12/10/2017 9:01:01 AM","Created":"12/10/2017 9:01:01 AM","ItemChildCount":0,"FolderChildCount":0,"fileType":"pptx","FileRef":"/sites/khDev/Shared Documents/a - Copy (2).pptx","FSObjType":0,"ServerRedirectedEmbedUri":"https://iglobe.sharepoint.com/sites/khDev/_layouts/15/WopiFrame.aspx?sourcedoc={b8748218-79ed-4606-8f69-865425f7b22f}\u0026action=interactivepreview"},{"ID":"38","FileLeafRef":"a - Copy.pptx","Modified":"12/10/2017 9:01:05 AM","Created":"12/10/2017 9:01:05 AM","ItemChildCount":0,"FolderChildCount":0,"fileType":"pptx","FileRef":"/sites/khDev/Shared Documents/a - Copy.pptx","FSObjType":0,"ServerRedirectedEmbedUri":"https://iglobe.sharepoint.com/sites/khDev/_layouts/15/WopiFrame.aspx?sourcedoc={015cd13b-eba8-41fd-a5d5-f2bb240a7a85}\u0026action=interactivepreview"},{"ID":"39","FileLeafRef":"a.pptx","Modified":"12/10/2017 9:01:09 AM","Created":"12/10/2017 9:01:09 AM","ItemChildCount":0,"FolderChildCount":0,"fileType":"pptx","FileRef":"/sites/khDev/Shared Documents/a.pptx","FSObjType":0,"ServerRedirectedEmbedUri":"https://iglobe.sharepoint.com/sites/khDev/_layouts/15/WopiFrame.aspx?sourcedoc={45be062d-88e4-4d1a-863d-995c4a29880c}\u0026action=interactivepreview"}]},
      documentList: [],
      profilename: '',
      menudata: [],
      loading: false,
      selectKey: ''

    } //state end

    this.menuInfo = this.menuInfo.bind(this);
    this.menuClick = this.menuClick.bind(this);
    this.menuLinkClick = this.menuLinkClick.bind(this);
  }

  componentWillMount() {
  }

  componentDidMount() {
    //this.menuInfo();

    this.setState({ documentList: this.state.userDocuments["document"] });

    debugger;
    this.state.userDocuments["document"].forEach(element => {
      let menuitem = null;
      let onefolder = {
        name: element.FileLeafRef,
        title: element.FileRef,
        key: element.ID,
        icon: 'FabricFolderFill'
      };
      let multiplefolder = {
        name: element.FileLeafRef,
        title: element.FileRef,
        key: element.ID,
        isExpanded: false,
        links: [
          {
            name: 'loading....',
            icon: 'FabricFolderFill'
          }
        ]
      };
      menuitem = (element.FSObjType === 1) && (element.FolderChildCount === 0) && onefolder;
      if (!menuitem) {
        menuitem = (element.FolderChildCount >= 1) && multiplefolder;
      }
      menuitem && this.state.menudata.push(menuitem);

    });


    this.setState({ menudata: this.state.menudata });
    this.state.menudata.length > 0 &&
      this.setState({ selectKey: this.state.menudata[0].key });
    this.setState({ loading: false });




    console.log(this.state.menudata);

  }

  componentWillUpdate() {
  }

  componentWillUnMount() {
  }

  menuLinkClick(e: React.MouseEvent<HTMLElement>, item: INavLink) {
    this.setState({ selectKey: item.key });

    if (!item.isExpanded) {
      (item.links.length === 1) && (item.links[0].name === 'loading....') &&
        item.links.pop();
      if (item.links.length === 0) {

        this.setState({ loading: true });

        let apiroot = "MenuBarTree/FoldersInfo?path=" + item.title;
        const url = this.state.urlroot + apiroot;
        axios.get(url)
          .then(response => {
            this.setState({ documentList: response.data });

            debugger;
            this.state.documentList.forEach(element => {
              let menuitem = null;
              let onefolder = {
                name: element.FileLeafRef,
                title: element.FileRef,
                key: element.ID,
                icon: 'FabricFolderFill'
              };
              let multiplefolder = {
                name: element.FileLeafRef,
                title: element.FileRef,
                key: element.ID,
                isExpanded: false,
                links: [
                  {
                    name: 'loading....',
                    icon: 'FabricFolderFill'
                  }
                ]
              };
              menuitem = (element.FSObjType === 1) && (element.FolderChildCount === 0) && onefolder;
              if (!menuitem) {
                menuitem = (element.FolderChildCount >= 1) && multiplefolder;
              }
              menuitem && item.links.push(menuitem);

            });

            this.setState({ loading: false });

          })
          .catch(error => {
            this.setState({ loading: false });
          });


      }
    }
  }

  menuClick(e: React.MouseEvent<HTMLElement>, item: INavLink) {
    this.setState({ selectKey: item.key });
    debugger;
  }

  menuInfo() {
    this.setState({ loading: true });
    let apiroot = "Home/MenuInfo";
    const url = this.state.urlroot + apiroot;
    axios.get(url)
      .then(response => {
        this.setState({ userDocuments: response.data });

        if (this.state.userDocuments !== '') {
          debugger;
          this.setState({ profilename: this.state.userDocuments["currentUser"] });
          this.setState({ documentList: this.state.userDocuments["document"] });
        }









      })
      .catch(error => {
        this.setState({ loading: false });
      });

  }



  render() {
    debugger;
    return (
      <div className="ms-Grid">

        <div className="ms-Grid-row">
          <div className="App-header">

            <div className="ms-Grid-col ms-md8">
              <img className="headerpic" src="./favicon.ico" />
              <span className="App-title">
                <span className="header1">Office 365</span>
                <span className="header2">Sharepoint</span>

              </span>
            </div>

            <div className="ms-Grid-col ms-md4">

              <span className="header3">

                <Icon iconName='UserFollowed' className="profilelogo" />
                {
                  (() => {
                    switch (this.state.profilename) {
                      case "": return "No Name";
                      default: return this.state.profilename;
                    }
                  })
                    ()}
              </span>
            </div>

          </div>
        </div>


        <div className="ms-Grid-row">
          <BarLoader
            color={'#0072c6'}
            width={500}
            height={2}
            loading={this.state.loading}
          />
          <div className="searchrow">

            <div className="ms-Grid-col ms-md3">
              <img className="bandlogo" src="./favicon.ico" />
            </div>

            <div className="ms-Grid-col ms-md5">
              <SearchBox
                className="simplesearch"
                labelText="Seach a documents here ...."
              />
            </div>


            <div className="ms-Grid-col ms-md4">
              <DefaultButton
                className="advancesearch"
                iconProps={{ iconName: 'Settings' }}>
                ADVANCED SEARCH
                </DefaultButton>
              <Icon iconName='CollapseMenu' className='menuicon' />
            </div>

          </div>

        </div>

        <div className="ms-Grid-row">
          <div className="ms-Grid-col ms-md3">

            <div className="menu">

              <Nav
                className="menunav"
                onLinkClick={this.menuClick}
                onLinkExpandClick={this.menuLinkClick}
                groups={
                  [{
                    links:
                      this.state.menudata
                  }]
                }
                parentId
                expandedStateText={'expanded'}
                collapsedStateText={'collapsed'}
                selectedKey={this.state.selectKey}
              />

            </div>

          </div>

          <div className="ms-Grid-col ms-md9">
             <Test/>
            {
              this.state.documentList.length > 0 &&
              <ShowDocuments documentsRender={this.state.documentList} />
            }

          </div>

        </div>


      </div>
    );
  }
}





export default App;
