import * as React from "react";
import {
    Provider,
    Flex,
    Text,
    Button,
    Header,
    ThemePrepared,
    themes,
    Input
} from "@fluentui/react-northstar";
import TeamsBaseComponent, { ITeamsBaseComponentState } from "msteams-react-base-component";
import * as microsoftTeams from "@microsoft/teams-js";
/**
 * State for the launchTab React component
 */
export interface ILaunchTabState extends ITeamsBaseComponentState {
    entityId?: string;
    teamsTheme: ThemePrepared;
}

/**
 * Properties for the launchTab React component
 */
export interface ILaunchTabProps {

}

/**
 * Implementation of the launchTab content page
 */
export class LaunchTab extends TeamsBaseComponent<ILaunchTabProps, ILaunchTabState> {

    public async componentWillMount() {

        this.updateComponentTheme(this.getQueryVariable("theme"));

        if (await this.inTeams()) {
            microsoftTeams.initialize();
            microsoftTeams.registerOnThemeChangeHandler(this.updateComponentTheme);
            microsoftTeams.getContext((context) => {
                microsoftTeams.appInitialization.notifySuccess();
                this.setState({
                    entityId: context.entityId
                });
                this.updateTheme(context.theme);
            });
        } else {
            this.setState({
                entityId: "This is not hosted in Microsoft Teams"
            });
        }
    }

    public render() {
        return (
            <Provider theme={this.state.teamsTheme}>
                <Flex column gap="gap.smaller">
                    <header>App Launch POC</header>
                    <br />
                    <Button content="Launch the app" primary onClick={this.onRedirect}></Button>
                </Flex>
            </Provider>
        );
    }

    private onRedirect = (event: React.MouseEvent<HTMLButtonElement>): void => {
        const taskModuleInfo = {
          title: "Redirect",
          url: this.appRoot() + `/launchTab/redirectTaskmodule.html`,
          width: 10,
          height: 10
        };
        microsoftTeams.tasks.startTask(taskModuleInfo);
      }

    private appRoot(): string {
        if (typeof window === "undefined") {
          return "https://{{HOSTNAME}}";
        } else {
          return window.location.protocol + "//" + window.location.host;
        }
      }

    private updateComponentTheme = (teamsTheme: string = "default"): void => {
        let theme: ThemePrepared;

        switch (teamsTheme) {
            case "default":
                theme = themes.teams;
                break;
            case "dark":
                theme = themes.teamsDark;
                break;
            case "contrast":
                theme = themes.teamsHighContrast;
                break;
            default:
                theme = themes.teams;
                break;
        }
        // update the state
        this.setState(Object.assign({}, this.state, {
            teamsTheme: theme
        }));
    }
}
