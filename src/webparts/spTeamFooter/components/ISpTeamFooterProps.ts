import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient } from '@microsoft/sp-http';

export interface ISpTeamFooterProps {
  listId: string;
  centerDirector: string;
  centerDirectorData: any;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
  httpClient: SPHttpClient;
  siteUrl: string;
}

export interface ICenterManager {
  id: number;
  title: string;
  email: string;
  picture?: string;
  department?: string;
  jobTitle?: string;
  location?: string;
}

export interface ITeamData {
  id: number;
  teamName: string;
  teamDescription: string;
  locations: string[];
  teamLeaders: any[];
  techLeaders: any[];
  centerManager: any;
}
