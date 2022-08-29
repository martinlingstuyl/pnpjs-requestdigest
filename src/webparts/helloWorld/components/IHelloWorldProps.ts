import { ServiceScope } from "@microsoft/sp-core-library";

export interface IHelloWorldProps {
  siteUrl: string;
  serviceScope: ServiceScope;
  formDigestValue: string;
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
}
