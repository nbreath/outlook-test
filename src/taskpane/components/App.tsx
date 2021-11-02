import * as React from "react";
import { DefaultButton } from "@fluentui/react";
import Header from "./Header";
import HeroList, { HeroListItem } from "./HeroList";
import Progress from "./Progress";

/* global require */

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {
  result: {};
  listItems: HeroListItem[];
}

export default class App extends React.Component<AppProps, AppState> {
  constructor(props, context) {
    super(props, context);
    this.state = {
      result: {},
      listItems: [],
    };
  }

  componentDidMount() {
    this.setState({
      result: {},
      listItems: [
        {
          icon: "Ribbon",
          primaryText: "Achieve more with Office integration",
        },
        {
          icon: "Unlock",
          primaryText: "Unlock features and functionality",
        },
        {
          icon: "Design",
          primaryText: "Create and visualize like a pro",
        },
      ],
    });
  }

  click = async () => {
    /**
     * Insert your Outlook code here
     */
    try {
      Office.context.mailbox.item.to.getAsync((result) => {
        console.log({result});
        this.setState(prevState => ({...prevState, result}))
      });
    } catch(exception) {
      console.log({exception});
      this.setState(prevState => ({...prevState, result: exception}))
    }
  };

  render() {
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress
          title={title}
          logo={require("./../../../assets/logo-filled.png")}
          message="Please sideload your addin to see app body."
        />
      );
    }

    return (
      <div className="ms-welcome">
        <Header logo={require("./../../../assets/logo-filled.png")} title={this.props.title} message="Welcome" />
        <HeroList message="Discover what Office Add-ins can do for you today!" items={[]}>
          <p className="ms-font-l">
            Modify the source files, then click <b>Execute item.to.getAsync</b>.
          </p>
          <p className="ms-font-l">
            Result of getAsync:
          </p>
          <p className="ms-font-l" style={{wordBreak: 'break-all'}}>
            {JSON.stringify(this.state.result)}
          </p>
          <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }} onClick={this.click}>
            Execute item.to.getAsync
          </DefaultButton>
        </HeroList>
      </div>
    );
  }
}
