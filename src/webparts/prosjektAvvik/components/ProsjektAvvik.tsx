import * as React from 'react';
import styles from './ProsjektAvvik.module.scss';
import { IItemResponse } from '../services/IPnPService';
import { SPHttpClient } from "@microsoft/sp-http";
import SPHttpClientResponse from "@microsoft/sp-http/lib/spHttpClient/SPHttpClientResponse";
import { Spinner, SpinnerSize, MessageBar, MessageBarType, Dialog, DialogType, Button, Label, Panel, PanelType, TextField, CommandBar, CommandBarButton,ICommandBarProps } from 'office-ui-fabric-react';
import { ListView, IViewField, SelectionMode } from "@pnp/spfx-controls-react/lib/ListView";
import { IProsjektAvvikProps, IProsjektAvviksItems, IProsjektAvvikState } from './IProsjektAvvikProps';
import { escape } from '@microsoft/sp-lodash-subset';
import pnp from 'sp-pnp-js';
import PnPService from '../services/PnPService';
import * as moment from 'moment';
import * as uuidv4 from 'uuid/v4';
import * as strings from 'ProsjektAvvikWebPartStrings';

export default class ProsjektAvvik extends React.Component<IProsjektAvvikProps, IProsjektAvvikState> {
  private _results: any[] = [];
  private _compId: string = "";
  private _PnPService: PnPService;

  constructor(props: IProsjektAvvikProps, state: IProsjektAvvikState) {
    super(props);

    // initialize the service
    this._PnPService = new PnPService(props.context);

    // specify a unique id for the component
    this._compId = "avvik" + uuidv4();

    // initialize the current component state
    this.state = {
      loading: true,
      projectNumber: "",
      reactListItems: [],
      error: "",
      showPanel: false,
      showError: false,
      panelInfo: {
        AnsvarForOppfolging: "",
        Avvik: "-",
        Arsak: "-",
        Beskrivelse: "-",
        Title: "-",
        Kategori: "-",
        Created: "-",
        Opprettet: "-",
        BakenforeliggendeArsak: "-",
        RelatertTilKundeLeverandor: "-",
        RelatertTilProsjekt: "-",
        AvvikType: "-",
        OkonomiskKonsekvens: "-",
        SpesifisertType: "-"
      }
    };


  }

  public componentDidMount(): void {
    this._processResults();
  }

  private _processResults() {
    let pNumber: string;
    let orderBy: string = "Created";
    let fields = "*,Kategori/Title,BakenforeliggendeArsak/Title,AvvikType/Title,Arsak/Title,SpesifisertType/Title,AnsvarForOppfolging/Title";
    let expand: string = "Kategori,BakenforeliggendeArsak,AvvikType,Arsak,SpesifisertType,AnsvarForOppfolging";

    this._PnPService._getProjectInfo("Prosjektinformasjon").then( resp => {

      if( typeof resp.projectID == "undefined" ) {
        pNumber = "0";
        this.setState({
          error: "Kunne ikke finne prosjektnummeret på denne siten, sjekk prosjektinformasjonslisten"
        });
      } else {
        pNumber = resp.projectID;
      }
      let query: string = "RelatertTilProsjekt eq '" + pNumber + "'"; // todo: retrieve project number from current web
      this._PnPService.get(query,
        orderBy,
        fields,
        expand,
        this.props.maxResults
      ).then((itemsResponse: IItemResponse) => {
        const itemValues: any = {
          wpTitle: this.props.title,
          pageCtx: this.props.context.pageContext,
          items: itemsResponse.results,
          totalResults: itemsResponse.totalResults
        };

        let itemsHtml: string = "";
        let renderItems: IProsjektAvviksItems[] = [];

        for (let item of itemValues.items) {
          let avvik: IProsjektAvviksItems = {};
          for (let prop in this.state.panelInfo) {
            if(item.hasOwnProperty(prop) && item[prop] !== null) {
              // lookup columns
              if (item[prop].hasOwnProperty("Title")) {
                avvik[prop] = item[prop].Title;
              } else {
                // render date correctly
                if (prop == "Created") {
                  avvik.Opprettet = moment(item.Created).format("DD/MM/YY HH:mm");
                } else {
                  avvik[prop] = item[prop] || "-";
                }
              }
            }
          }

          renderItems.push(avvik);
        }

        this.setState({
          loading: false,
          projectNumber: pNumber,
          reactListItems: renderItems
        });

      }).catch( (error: any) => {
        this.setState({
          error: error.toString()
        });
      });
    }).catch( (error: any) => {
      this.setState({
        error: error.toString()
      });
    });
  }

  public render(): React.ReactElement<IProsjektAvvikProps> {
    let view = <Spinner size={SpinnerSize.large} label='Laster avvik' />;
    let powerAppsLink: string = "https://web.powerapps.com/webplayer/app?web=1&projectnr=" + this.state.projectNumber + "&appId=%2fproviders%2fMicrosoft.PowerApps%2fapps%2f15cc36c1-3ae8-4779-b030-dde92f78b935";

    if (this.props.powerapplink) {
      let projectToken: string = "projectnr=";
      let locationOfToken: number = this.props.powerapplink.indexOf(projectToken)+projectToken.length;
      powerAppsLink = this.spliceString(this.props.powerapplink,locationOfToken,0,this.state.projectNumber);
    }

    const viewFields: IViewField[] = [
      {
        name: 'Title',
        displayName: 'Avvik',
        maxWidth: 220,
        sorting: true
      },
      {
        name: 'Kategori',
        maxWidth: 180,
        sorting: true
      },
      {
        name: 'Opprettet',
        maxWidth: 150
      }
    ];

    const commandBarItems: ICommandBarProps = {
      items: [{
        key: "nyttAvvik",
        name: "Nytt avvik",
        iconProps: {
          iconName: "PowerApps2Logo"
        },
        href: powerAppsLink
      }]
    };

    if ( !this.state.loading && this.state.reactListItems.length == 0 && this.state.error == "" ) {
      return( <div>
        <CommandBar
        isSearchBoxVisible={ false }
        items={commandBarItems.items}
      />
        <MessageBar messageBarType={MessageBarType.info}>
          <span>Ingen avvik registrert p&aring; dette prosjektet</span>
        </MessageBar>
      </div>
      );
    }

    if (!this.state.loading && this.state.reactListItems && this.state.error == "") {
      view = <div>
          <CommandBar
            isSearchBoxVisible={ false }
            items={commandBarItems.items}
          />
          <ListView
          items={this.state.reactListItems}
          compact={true}
          viewFields={viewFields}
          selectionMode={SelectionMode.single}
          selection={ (e) => this._getSelection(e,this) } />
          <Panel
            isBlocking={ false }
            isOpen={ this.state.showPanel }
            type={ PanelType.smallFixedFar }
            // tslint:disable-next-line:jsx-no-lambda
            onDismiss={ () => this.setState({
              showPanel: false
              }) }
            headerText='Avvik'
          >
            <TextField
              label= "Tittel"
              placeholder={this.state.panelInfo.Title}
              disabled={ true }
            />
            <TextField
              label= "Kategori"
              placeholder={this.state.panelInfo.Kategori}
              disabled={ true }
            />
            <TextField
              label= "Opprettet"
              placeholder={this.state.panelInfo.Opprettet}
              disabled={ true }
            />
            <TextField
              label= "Beskrivelse"
              multiline
              rows={ 5 }
              placeholder={this.state.panelInfo.Beskrivelse}
              disabled={ true }
            />
            <TextField
              label= "Ansvarlig"
              placeholder={this.state.panelInfo.AnsvarForOppfolging}
              disabled={ true }
            />
            <TextField
              label= "Kunde / Leverandør"
              placeholder={this.state.panelInfo.RelatertTilKundeLeverandor}
              disabled={ true }
            />
            <TextField
              label= "Avvikstype"
              placeholder={this.state.panelInfo.AvvikType}
              disabled={ true }
            />
            <TextField
              label= "Økonomisk konsekvens"
              placeholder={this.state.panelInfo.OkonomiskKonsekvens}
              disabled={ true }
            />
            <TextField
              label= "Årsak"
              placeholder={this.state.panelInfo.Arsak}
              disabled={ true }
            />
            <TextField
              label= "Bakenforliggende årsak"
              placeholder={this.state.panelInfo.BakenforeliggendeArsak}
              disabled={ true }
            />
            <TextField
              label= "Spesifisert type"
              placeholder={this.state.panelInfo.SpesifisertType}
              disabled={ true }
            />
          </Panel>
        </div>;
    }
    if (this.state.error !== "") {
      return (
          <MessageBar className={styles.error} messageBarType={MessageBarType.error}>
              <span>Sorry, something went wrong</span>
              {
                  (() => {
                      if (this.state.showError) {
                          return (
                              <div>
                                  <p>
                                      <a href="javascript:;" onClick={this._toggleError.bind(this)} className="ms-fontColor-neutralPrimary ms-font-m">
                                      <i className={`ms-Icon ms-Icon--ChevronUp ${styles.icon}`} aria-hidden="true"></i> Hide the error message</a>
                                  </p>
                                  <p className="ms-font-m">{this.state.error}</p>
                              </div>
                          );
                      } else {
                          return (
                              <p>
                                  <a href="javascript:;" onClick={this._toggleError.bind(this)} className="ms-fontColor-neutralPrimary ms-font-m">
                                  <i className={`ms-Icon ms-Icon--ChevronDown ${styles.icon}`} aria-hidden="true"></i> Show the error message</a>
                              </p>
                          );
                      }
                  })()
              }
          </MessageBar>
      );
    }



    return (
      <div id={this._compId} className={styles.prosjektAvvik}>
          {view}
          <Dialog isOpen={this.state.showScriptDialog} type={DialogType.normal} onDismiss={this._toggleDialog.bind(this)} title={strings.ScriptsDialogHeader} subText={strings.ScriptsDialogSubText}></Dialog>
      </div>
  );
  }

  public _getSelection(item: any, ctx: any) {
    if (item.length !== 0) {
      ctx.setState({
        showPanel: true,
        panelInfo: item[0]
      });
    }
  }

  private spliceString (stringToSplice: string, idx: number, rem: number, str: string) {
    return stringToSplice.slice(0, idx) + str + stringToSplice.slice(idx + Math.abs(rem));
  }

  private _toggleError() {
    this.setState({
        showError: !this.state.showError
    });
  }

  /**
   * Toggle the script dialog visibility
   */
  private _toggleDialog() {
      this.setState({
          showScriptDialog: !this.state.showScriptDialog
      });
  }

}
