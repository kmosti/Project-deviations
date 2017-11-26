import { IWebPartContext } from "@microsoft/sp-webpart-base";
import { IViewField } from "@pnp/spfx-controls-react/lib/ListView";
import * as moment from 'moment';
import { ILinkItemProps } from './ILinkItemProps';

export interface IProsjektAvvikProps {
  title: string;
  maxResults: number;
  powerapplink: string;
  context: IWebPartContext;
  linkItems?: Array<ILinkItemProps>;
  editItem?: Function;
}

export interface IProsjektAvvikPanelInfo extends IProsjektAvviksItems {

}

export interface IProsjektAvvikState {
  loading?: boolean;
  projectNumber?: string;
  reactListItems?: IProsjektAvviksItems[];
  viewFields?: IViewField ;
  error?: string;
  showError?: boolean;
  showPanel?: boolean;
  showScriptDialog?: boolean;
  panelInfo?: IProsjektAvviksItems;
}

export interface IProsjektInformasjon {
  Title?: string;
  projectID?: string;
}

export interface IProsjektAvviksItems {
  AnsvarForOppfolgingId?: number;
  AnsvarForOppfolging?: string;
  AnsvarForOppfolgingStringId?: string;
  ArsakId?: number;
  Arsak?: string;
  Attachments?: boolean;
  AuthorId?: number;
  Avvik?: string;
  AvvikTypeId?: number;
  AvvikType?: string;
  BakenforeliggendeArsakId?: number;
  BakenforeliggendeArsak?: string;
  Beskrivelse?: string;
  ComplianceAssetId?: number;
  ContentTypeId?: string;
  Created?: any;
  DatoForHendelse?: string;
  EditorId?: number;
  EstimertKostnad?: number;
  FileSystemObjectType?: number;
  GUID?: string;
  ID?: number;
  Id?: number;
  KategoriId?: number;
  Kategori?: string;
  Konsekvens?: string;
  Modified?: string;
  Opprettet?: string;
  OkonomiskKonsekvens?: string;
  RelatertTilKundeLeverandor?: string;
  RelatertTilProsjekt?: string;
  SpesifisertTypeId?: string;
  SpesifisertType?: string;
  TiltakForbedringsforslag?: string;
  Title?: string;
}
