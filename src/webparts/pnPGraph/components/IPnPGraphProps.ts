import { SPHttpClient} from '@microsoft/sp-http';
import { IPersona } from 'office-ui-fabric-react';

export interface IPnPGraphProps {
  description: string;
  myhttp : SPHttpClient;  
  mysite: string;
  siteurl: string;
  me:  {email:string,displayname:string,phone:string,firstname:string};
}
