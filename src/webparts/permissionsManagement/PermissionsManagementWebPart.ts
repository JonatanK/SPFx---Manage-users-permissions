import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './PermissionsManagementWebPart.module.scss';
import * as strings from 'PermissionsManagementWebPartStrings';

import pnp from 'sp-pnp-js';
import { Web } from 'sp-pnp-js';


export interface IPermissionsManagementWebPartProps {
  description: string;
}

export default class PermissionsManagementWebPartWebPart extends BaseClientSideWebPart<IPermissionsManagementWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.permissionsManagement }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <span class="${ styles.title }">Gerenciamento de Permissões SharePoint</span>
              <p class="${ styles.subTitle }">Verifique, adicione e remova usuários de grupos de segurança.</p>
              <div id="three-view">
              </div>
            </div>
          </div>
        </div>
      </div>`;
      this._renderSiteCollections();
      this._setButtonClicks();

  }

  private _setButtonClicks(){
    let elements: HTMLCollectionOf<Element> = document.getElementsByClassName('expand-site');
    console.log("Elementos", elements)
    for (let i = 0; i < elements.length; i++) {
        elements[i].parentElement.setAttribute('id', `siteRow-${i}`);
        elements[i].setAttribute('siteIndex', i.toString());
        elements[i].addEventListener('click', () => this._getPermissionGroups(elements[i].getAttribute('siteUrl'), i), true);
    }
  }


  private _renderSiteCollections(){
    let html: string = '';

    html += `<ul>`;
    let sites: string[] = this.properties.description.split(';')
    sites.forEach((item: string) => {
      if(item != ''){
        html += `<li>${item} - <a class="expand-site" siteUrl="${item}">expandir</a></li>`
      }
    })
    html += `</ul>`;

    this.domElement.querySelector('#three-view').innerHTML = html;

  }

  private _getPermissionGroups(url:string, position:number){
    let html: string = '';
    let web = new Web(url);

    web.siteGroups.get()
    .then((groups) => {
      console.log(groups)
      html += '<ul>'
      groups.forEach((g) => {
        html += `<li id="group-users-list-${g.Id}">${g.Title} - <a class="manage-users" groupId="${g.Id}" siteUrl="${url}">gerenciar usuários</a></li>`
      })
      html += '</ul>'

      this.domElement.querySelector(`#siteRow-${position}`).innerHTML += html

      let elements: HTMLCollectionOf<Element> = document.getElementsByClassName('manage-users')
      for (let i = 0; i < elements.length; i++) {
          elements[i].addEventListener('click', () => this._manageUsers(elements[i].getAttribute('siteUrl'), parseInt(elements[i].getAttribute('groupId'))), true);
      }

    })
    .catch((error) => {
      console.error(error)
    })
  }

  private _manageUsers(url:string, groupId:number){
    let web = new Web(url)
    web.siteGroups.getById(groupId).users.get()
    .then((users) => {
      console.log(users)
      let html = `<ul>`
      users.forEach(user => {
          html += `<li>${user.Title}</li>`
      });
      html += `</ul>`

      this.domElement.querySelector(`#group-users-list-${groupId}`).innerHTML += html
    })
    .catch((error) => {
      console.error(error);
    })
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: 'Configurações da webpart'
          },
          groups: [
            {
              groupName: 'Site collections',
              groupFields: [
                PropertyPaneTextField('description', {
                  label: 'Sites collections',
                  description: 'Informe os site collections que deverão ser gerenciados separados por ponto-e-virgula (;)',
                  multiline: true,
                  rows: 10
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
