import React, { Component } from 'react';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import './ShowDocuments.css';
export class ShowDocuments extends React.Component {

  constructor(props) {
    debugger;
    super(props);
    this.state = {
      data: props.documentsRender
    }

  }

  render() {
    debugger;
    return (
      <div className="ms-Grid-row">
        <div className="ms-Grid-col ms-md9">
          <table className="documents">
            <thead>
              <tr>
                <th>Name</th>
                <th>Created</th>
                <th>Modified</th>
              </tr>
            </thead>
            <tbody>
              {
                this.state.data.map((element, i) =>
                  <Content key={i} documentRow={element} />
                )
              }
            </tbody>
          </table>
        </div>
        <div className="ms-Grid-col ms-md3"></div>
      </div>


    );
  }

}



class Content extends React.Component {

  constructor(props) {
    debugger;
    super(props);
  }
  render() {
    return (
      <tr>
        <td>{this.props.documentRow.FileLeafRef}</td>
        <td>{this.props.documentRow.Created}</td>
        <td>{this.props.documentRow.Modified}</td>
      </tr>

    );
  }
}