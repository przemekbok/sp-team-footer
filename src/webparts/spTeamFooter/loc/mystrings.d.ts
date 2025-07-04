declare interface ISpTeamFooterWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppLocalEnvironmentOffice: string;
  AppLocalEnvironmentOutlook: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  AppOfficeEnvironment: string;
  AppOutlookEnvironment: string;
  UnknownEnvironment: string;
  ListFieldLabel: string;
  CenterDirectorFieldLabel: string;
  CreateNewListLabel: string;
  ViewSelectedListLabel: string;
}

declare module 'SpTeamFooterWebPartStrings' {
  const strings: ISpTeamFooterWebPartStrings;
  export = strings;
}
