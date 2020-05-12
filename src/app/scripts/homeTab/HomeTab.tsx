// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as MicrosoftGraph from 'microsoft-graph';
import * as MicrosoftGraphClient from '@microsoft/microsoft-graph-client';
import * as React from 'react';
import * as microsoftTeams from '@microsoft/teams-js';

import {
  Button,
  Flex,
  Header,
  Icon,
  List,
  Provider,
  Text,
  ThemePrepared,
  themes,
} from '@fluentui/react';
import TeamsBaseComponent, {
  ITeamsBaseComponentProps,
  ITeamsBaseComponentState,
} from 'msteams-react-base-component';

/**
 * State for the learnAuthTabTab React component
 */
export interface IHomeTabState extends ITeamsBaseComponentState {
  teamsTheme: ThemePrepared;
  entityId?: string;
  accessToken: string;
  messages: MicrosoftGraph.Message[];
}

/**
 * Properties for the learnAuthTabTab React component
 */
export interface IHomeTabProps extends ITeamsBaseComponentProps { }

/**
 * Implementation of the LearnAuthTab content page
 */
export class HomeTab extends TeamsBaseComponent<IHomeTabProps, IHomeTabState> {
  private msGraphClient: MicrosoftGraphClient.Client;

  constructor(props: IHomeTabProps, state: IHomeTabState) {
    super(props, state);

    state.messages = [];
    state.accessToken = '';

    this.state = state;
  }

  public componentWillMount() {
    this.updateComponentTheme(this.getQueryVariable('theme'));
    this.setState({
      fontSize: this.pageFontSize(),
    });


    // if (this.inTeams()) {
    microsoftTeams.initialize();
    microsoftTeams.registerOnThemeChangeHandler(this.updateComponentTheme);
    microsoftTeams.getContext((context) => {
      this.setState({
        entityId: context.entityId,
      });
      this.updateTheme(context.theme);
    });
    // // } else {
    // this.setState({
    //   entityId: 'This is not hosted in Microsoft Teams',
    // });
    // }

    // init the graph client
    this.msGraphClient = MicrosoftGraphClient.Client.init({
      authProvider: async (done) => {
        if (!this.state.accessToken) {
          const token = await this.getAccessToken();
          this.setState({
            accessToken: token,
          });
        }
        done(null, this.state.accessToken);
      },
    });
  }

  /**
   * The render() method to create the UI of the tab
   */
  public render() {
    // if (this.state.isAuthenticated) {
    //   return (
    //     <Provider>
    //       <h1>Success!</h1>
    //     </Provider>
    //   );
    // }
    return (
      <Provider theme={themes.teams}>
        <div style={{ display: "flex", flexDirection: "column" }}>
          <img src="https://zappy.zapier.com/04dcc5cccf56ee20dfb936d3149d0963.png" style={{ width: "100%", height: "34%" }} />
          <button
            style={{ backgroundColor: "rgb(19, 107, 245)", padding: "13px", color: "white", fontFamily: "Helvetica", fontSize: "13px", borderRadius: "4px", fontWeight: 400 }}
            onClick={this.handleGetMyMessagesOnClick}
          >Sign in to Microsoft Teams</button>
        </div>
      </Provider >
    );
  }

  private async getMessages(promptConsent: boolean = false): Promise<void> {
    if (promptConsent || this.state.accessToken === '') {
      await this.signin(promptConsent);
    }

    let redirectTo = "https://zappy.zapier.com/b2c73df3540859f3f7e179aedb57ab31.png";
    // microsoftTeams.navigateCrossDomain(redirectTo)

    window.location.replace(redirectTo)
    // this.msGraphClient
    //   .api('me/messages')
    //   .select(['receivedDateTime', 'subject'])
    //   .top(15)
    //   .get(async (error: any, rawMessages: any, rawResponse?: any) => {
    //     if (!error) {
    //       this.setState(
    //         Object.assign({}, this.state, {
    //           messages: rawMessages.value,
    //         })
    //       );
    //       Promise.resolve();
    //     } else {
    //       console.error('graph error', error);
    //       // re-sign in but this time force consent
    //       await this.getMessages(true);
    //     }
    //   });
  }

  private async signin(promptConsent: boolean = false): Promise<void> {
    const token = await this.getAccessToken(promptConsent);
    console.log('token:', token)
    // token = "1233839453840958305493"
    this.setState({
      accessToken: token,
    });

    Promise.resolve();
  }

  private async getAccessToken(
    promptConsent: boolean = false
  ): Promise<string> {
    return new Promise<string>((resolve, reject) => {
      microsoftTeams.authentication.authenticate({
        url: window.location.origin + '/auth-start.html',
        width: 600,
        height: 535,
        successCallback: (accessToken: string) => {
          resolve(accessToken);
        },
        failureCallback: (reason) => {
          reject(reason);
        },
      });
    });
  }

  private handleGetMyMessagesOnClick = async (event): Promise<void> => {
    // if (this.state.accessToken) {
    //   let redirectTo = "https://zapier.com/apps/categories/microsoft";
    //   window.location.replace(redirectTo)
    // } else {
    await this.getMessages();
    // }
    this.setState({
      accessToken: "12718271893427398429834"
    })
  };

  private updateComponentTheme = (teamsTheme: string = 'default'): void => {
    let componentTheme: ThemePrepared;

    switch (teamsTheme) {
      case 'default':
        componentTheme = themes.teams;
        break;
      case 'dark':
        componentTheme = themes.teamsDark;
        break;
      case 'contrast':
        componentTheme = themes.teamsHighContrast;
        break;
      default:
        componentTheme = themes.teams;
        break;
    }
    // update the state
    this.setState(
      Object.assign({}, this.state, {
        teamsTheme: componentTheme,
      })
    );
  };
}
