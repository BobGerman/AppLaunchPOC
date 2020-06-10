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
 * State for the pocTabTab React component
 */
export interface IPocTabState extends ITeamsBaseComponentState {
    entityId?: string;
    teamsTheme: ThemePrepared;
    youTubeVideoId?: string;
}

/**
 * Properties for the pocTabTab React component
 */
export interface IPocTabProps {

}

/**
 * Implementation of the POCTab content page
 */
export class PocTab extends TeamsBaseComponent<IPocTabProps, IPocTabState> {

    public async componentWillMount() {

        this.updateComponentTheme(this.getQueryVariable("theme"));
        this.setState(Object.assign({}, this.state, {
            youTubeVideoId: "jugBQqE_2sM"
        }));

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
{
    /*
                    <Header>Task Module Demo</Header>
                    <Text>YouTube Video ID:</Text>
                    <Input value={this.state.youTubeVideoId} disabled></Input>
    */
}
                    <header>App Launch POC</header>
                    <br />
                    <Button content="Launch the app" primary onClick={this.onRedirect}></Button>
{
    /*
                    <Button content="Change Video ID" onClick={this.onChangeVideo}></Button>
                    <Button content="Show Video" primary onClick={this.onShowVideo}></Button>
                    <Text content="(C) Copyright Contoso" size="smallest"></Text>
    */
}
                </Flex>
            </Provider>
        );
    }

    private onRedirect = (event: React.MouseEvent<HTMLButtonElement>): void => {
        const taskModuleInfo = {
          title: "Redirect",
          url: this.appRoot() + `/pocTab/redirectTaskmodule.html`,
          width: 10,
          height: 10
        };
        microsoftTeams.tasks.startTask(taskModuleInfo);
      }

    private onShowVideo = (event: React.MouseEvent<HTMLButtonElement>): void => {
        const taskModuleInfo = {
          title: "YouTube Player",
          url: this.appRoot() + `/pocTab/taskmodule.html?vid=${this.state.youTubeVideoId}`,
          width: 1000,
          height: 700
        };
        microsoftTeams.tasks.startTask(taskModuleInfo);
      }

    private onChangeVideo = (event: React.MouseEvent<HTMLButtonElement>): void => {
        const taskModuleInfo = {
            title: "YouTube Video Selector",
            url: this.appRoot() + `/pocTab/selector.html?theme={theme}&vid=${this.state.youTubeVideoId}`,
            width: 350,
            height: 150
          };
          
          const submitHandler = (err: string, result: string): void => {
            this.setState(Object.assign({}, this.state, {
              youTubeVideoId: result
            }));
          };
          
          microsoftTeams.tasks.startTask(taskModuleInfo, submitHandler);
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
