import * as React from "react";
import * as microsoftTeams from "@microsoft/teams-js";

import { Flex, Header, Input, Provider } from "@fluentui/react";
import TeamsBaseComponent, { ITeamsBaseComponentProps, ITeamsBaseComponentState } from "msteams-react-base-component";

export interface IHomeTabConfigState extends ITeamsBaseComponentState {
    value: string;
}

export interface IHomeTabConfigProps extends ITeamsBaseComponentProps {

}

/**
 * Implementation of home configuration page
 */
export class HomeTabConfig extends TeamsBaseComponent<IHomeTabConfigProps, IHomeTabConfigState> {

    public componentWillMount() {
        this.updateTheme(this.getQueryVariable("theme"));

        if (this.inTeams()) {
            microsoftTeams.initialize();

            microsoftTeams.getContext((context: microsoftTeams.Context) => {
                this.setState({
                    value: context.entityId
                });
                this.updateTheme(context.theme);
                this.setValidityState(true);
            });

            microsoftTeams.settings.registerOnSaveHandler((saveEvent: microsoftTeams.settings.SaveEvent) => {
                // Calculate host dynamically to enable local debugging
                const host = "https://" + window.location.host;
                microsoftTeams.settings.setSettings({
                    contentUrl: host + "/homeTab/?data=",
                    websiteUrl: host + "/homeTab/?data=",
                    suggestedDisplayName: "home",
                    removeUrl: host + "/homeTab/remove.html",
                    entityId: this.state.value
                });
                saveEvent.notifySuccess();
            });
        } else {
        }
    }

    public render() {
        return null;
    }
}
