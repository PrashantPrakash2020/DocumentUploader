export interface IMeetingForm {
    attachements: string[];
    peoples: string[];
    fileInfo: any;
    status: boolean
    type: string;
    subtype: string;
    typech:string[];
    subtypech:string[];
    date:Date;
    product: string;
    amount: string;
  }
export type JsonPrimitive = string | number | boolean | null;
export interface IJsonMap extends Record<string, JsonPrimitive | IJsonArray | IJsonMap> { }
export interface IJsonArray extends Array<JsonPrimitive | IJsonArray | IJsonMap> { }
export type Json = JsonPrimitive | IJsonMap | IJsonArray;
export interface IDocumentLibraryView {
}
export interface ICustomListView {
  items: IItem[];
  loading: boolean;
  showDialog: boolean;
  folderName: string;
}
export interface IiframeView {
  folderName: string;
}
export interface IItem {
  [key: string]: string;
}
export interface ILibraryView {
  title: string;
  logo: string;
}