import * as React from 'react';
import styles from './MyTeamsSpFx.module.scss';
import { IMyTeamsSpFxProps } from './IMyTeamsSpFxProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { ServiceProvider } from '../../services/ServiceProvider';  

export interface IMyTeamsState {  
  
  myteams: any[];  
  selectedTeam: any;  
  teamChannels: any;  
  selectedChannel: any;  
}  
export default class MyTeamsSpFx extends React.Component<IMyTeamsSpFxProps, IMyTeamsState> {

  private serviceProvider;  
  private messageTextRef;  
  
  public constructor(props: IMyTeamsSpFxProps, state: IMyTeamsState) {  
    super(props);  
    this.serviceProvider = new ServiceProvider(this.props.context);  
    this.state = {  
      myteams: [],  
      selectedTeam: null,  
      selectedChannel: null,  
      teamChannels: []  
    };  
  }  
  
  private GetmyTeams() {  
  
    this.serviceProvider.  
      getmyTeams()  
      .then(  
        (result: any[]): void => {  
          console.log(result);  
          this.setState({ myteams: result });  
        }  
      )  
      .catch(error => {  
        console.log(error);  
      });  
  }  

  public render(): React.ReactElement<IMyTeamsSpFxProps> {
    return (  
      <React.Fragment>  
         <div className={ styles.myTeamsSpFx }>  
        <h1>Teams Operations Demo using Graph API</h1>  
          
        <div>  
          <button className={styles.buttons} onClick={() => this.GetmyTeams()}>Get My Teams</button>  
        </div>  
        { this.state.myteams.length>0 &&  
        <React.Fragment>  
         <h3>Below is list of your teams</h3>  
         <h4>Select any team and click on Get Channels</h4>  
        </React.Fragment>  
        }  
        {this.state.myteams.map(  
          (team: any, index: number) => (  
            <React.Fragment>  
              <input className={styles.radio} onClick={() => this.setState({ selectedTeam: team })} type="radio" id={team.id} name="myteams" value={team.id} />  
              <label >{team.displayName}</label><br />  
            </React.Fragment>  
          )  
        )  
        }  
  
        {this.state.selectedTeam &&  
        <React.Fragment>  
            <br></br>  
            <button  className={styles.buttons} onClick={() => this.getChannels()}>Get Channels</button>  
          </React.Fragment>  
        }  
        { this.state.teamChannels.length>0 &&  
        <React.Fragment>  
        <h3>Below is list of your channels for selected team : {this.state.selectedTeam.displayName}</h3>  
        <h4>Select any channel, enter message and click 'Send Message' to Post Message on MS teams</h4>  
        </React.Fragment>  
        }  
  
        {this.state.teamChannels.map(  
          (channel: any, index: number) => (  
            <React.Fragment>  
              <input className={styles.radio} onClick={() => this.setState({ selectedChannel: channel })} type="radio" id={channel.id} name="teamchannels" value={channel.id} />  
              <label >{channel.displayName}</label><br />  
            </React.Fragment>  
          )  
        )  
        }  
  
        {this.state.selectedChannel &&  
          <React.Fragment>  
          <br></br>  
          <div>  
            <input className={styles.textbox}  ref={(elm) => { this.messageTextRef = elm; }} type="text" id="message" name="message" />  
            <br></br>  
            <br></br>  
            <button  className={styles.buttons} onClick={() => this.sendMesssage()}>Send Message</button>  
              
          </div>  
          </React.Fragment>  
        }  
        </div>  
      </React.Fragment>  
    );  
  }


  private getChannels() {  
  
    this.serviceProvider.  
      getChannel(this.state.selectedTeam.id)  
      .then(  
        (result: any[]): void => {  
          console.log(result);  
          this.setState({ teamChannels: result });  
        }  
      )  
      .catch(error => {  
        console.log(error);  
      });  
  }  
  
  private sendMesssage() {   
    this.serviceProvider.  
      sendMessage(this.state.selectedTeam.id, this.state.selectedChannel.id, this.messageTextRef.value)  
      .then(  
        (result: any[]): void => {  
          alert("message posted sucessfully");  
        }  
      )  
      .catch(error => {  
        console.log(error);  
      });  
  }  
}  

