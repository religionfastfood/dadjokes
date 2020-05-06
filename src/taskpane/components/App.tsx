import * as React from "react";
import { Button, ButtonType } from "office-ui-fabric-react";
import Header from "./Header";
import HeroList, { HeroListItem } from "./HeroList";
import Progress from "./Progress";
import Axios from "axios";
import About from "./About";
/* global Button, Header, HeroList, HeroListItem, Progress */

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {
  listItems: HeroListItem[];
  joke: string;
  hasJoke: boolean;
  showAbout: boolean;
}

export default class App extends React.Component<AppProps, AppState> {
  constructor(props, context) {
    super(props, context);
    this.state = {
      listItems: [],
      joke: "",
      hasJoke: false,
      showAbout: false
    };
  }

  componentDidMount() {
    this.setState({
      listItems: [
        {
          icon: "PeopleBlock",
          primaryText: "No one will hate this!"
        },
        {
          icon: "ServerProcesses",
          primaryText: "Uses company resources!"
        },
        {
          icon: "UnknownCall",
          primaryText: "100% non-consensual"
        }
      ]
    });
  }

  click = async () => {
    Office.context.mailbox.displayNewMessageForm({
      toRecipients: [],
      subject: "Dad Joke!",
      htmlBody: this.state.joke
    });
  };

  getJokeHandler = async () => {
    Axios.get("https://icanhazdadjoke.com/", {
      headers: { Accept: "application/json", "User-Agent": "My Office Add-in (cmcenteemcdonald@onbase.onmicrosoft.com" }
    }).then(response => {
      this.setState({ joke: response.data.joke, hasJoke: true });
    });
  };

  aboutThisAddInHandler = () => {
    this.setState(
      { showAbout: !this.state.showAbout}
    );
  }

  render() {
    const { title, isOfficeInitialized } = this.props;

    const iconToggle = this.state.showAbout ? "ChevronLeft" : "Help";

    if (!isOfficeInitialized) {
      return (
        <Progress title={title} logo="assets/logo-filled.png" message="Please sideload your addin to see app body." />
      );
    }

    return (
      <div className="ms-welcome">
        <Header logo="assets/dad-clipart.png" title={this.props.title} message="DAD JOKES" />
        {this.state.hasJoke ? (
          <h3 className="ms-fontSize-24">&quot;{this.state.joke}&quot;</h3>
        ) : (
          <HeroList
            message="Share cringey, eye-rolling jokes with your coworkers!"
            items={this.state.listItems}
          ></HeroList>
        )}
        {this.state.showAbout ? <About></About> : null}
        <Button
          className="ms-welcome__action"
          buttonType={ButtonType.hero}
          iconProps={{ iconName: "Emoji" }}
          onClick={this.getJokeHandler}
        >
          {this.state.hasJoke ? "Get Another Joke!" : "Get a Joke!"}
        </Button>
        {this.state.hasJoke ? (
          <Button
            className="ms-welcome__action"
            buttonType={ButtonType.hero}
            iconProps={{ iconName: "MailForward" }}
            onClick={this.click}
          >
            Send to Contacts
          </Button>
        ) : null}
        <Button className="ms-welcome__action" buttonType={ButtonType.hero} iconProps={{ iconName: iconToggle }} onClick={this.aboutThisAddInHandler}>
          {this.state.showAbout ? "Back" : "About This Add-in"}
        </Button>
      </div>
    );
  }
}
