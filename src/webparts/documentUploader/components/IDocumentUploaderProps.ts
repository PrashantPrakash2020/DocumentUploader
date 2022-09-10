import { WebPartContext } from '@microsoft/sp-webpart-base';
import { DataService } from './DataService';
export interface IDocumentUploaderProps {
  description: string;
  context: WebPartContext;
}