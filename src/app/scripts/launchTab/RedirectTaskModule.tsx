import * as React from "react";
import { Provider, Flex, Text, Button, Header, ThemePrepared, themes, Input } from "@fluentui/react-northstar";
import TeamsBaseComponent, { ITeamsBaseComponentState } from "msteams-react-base-component";
import * as microsoftTeams from "@microsoft/teams-js";

export interface IRedirectTaskModuleState extends ITeamsBaseComponentState {
    teamsTheme: ThemePrepared;
}

export interface IRedirectTaskModuleProps { }

export class RedirectTaskModule extends TeamsBaseComponent<IRedirectTaskModuleProps, IRedirectTaskModuleState> {
    public componentWillMount(): void {

        if (this.inTeams()) {
            microsoftTeams.initialize();
            const searchParams = new URLSearchParams(window.location.search);
            const symbol = searchParams.get("symbol");
            let launchUrl = process.env.APP_URL;
            if (launchUrl && symbol) {
                launchUrl = launchUrl.replace("{$SYMBOL}", symbol);
            }
            window.open(launchUrl);
            microsoftTeams.tasks.submitTask();
        }
    }

    public render() {
        return (
            <Text size="medium">
                Run in Teams for correct operation
            </Text>
        );
    }
}
