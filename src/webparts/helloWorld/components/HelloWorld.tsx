import * as React from 'react';
import styles from './HelloWorld.module.scss';
import { IHelloWorldProps, IHelloWorldState } from './IHelloWorldProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { MSGraphClient } from '@microsoft/sp-client-preview';
import { SpinnerSize, Spinner, MessageBar, MessageBarType, Link, List } from 'office-ui-fabric-react';

export interface ICalendarEvent {
  Title: string;
  StartDate: any;
  EndDate: any;
  Url: string;
  Location?: any;
  SkypeUrl?: string;
  AllDayEvent?: boolean;
}

export default class HelloWorld extends React.Component<IHelloWorldProps, IHelloWorldState> {
  private client: MSGraphClient;
  constructor(props: IHelloWorldProps) {
    super(props);

    this.state = {
      items: []
    };
  }


  public componentDidMount() {
    let sem = null;
    if ((window as any).HelloWorldWP && (window as any).HelloWorldWP.MSGraphClientSemaphore) {
      sem = (window as any).HelloWorldWP.MSGraphClientSemaphore;
    }
    else {
      sem = require('./Semaphore.js')(1);
      (window as any).HelloWorldWP = {
        MSGraphClientSemaphore: sem
      };
    }

    sem.take(function () {
      // Get graph client from context, this was added in SPFx 1.4.1
      // With the removal of the popup (around April 10th) it's causing race conditions when there are multiple webparts on the page which call the MSGraphClient
      // To work around this we've added a semaphore so only one client can be fetched at a time
      // After the first fetch, the token is cached and other webparts can get it from there correctly.

        this.client = this.props.context.serviceScope.consume(MSGraphClient.serviceKey);
        // We need to fetch our items as well, the client will apparently only login when an actual call is made.
        // Calling _getOAuthToken ourselves doesn't work without rewriting everything, don't want to do that.
        this.getEvents(10).then(
          //resolve
          (apps: ICalendarEvent[]) => {
            sem.leave();
            this.setState({
              items: apps,
              isErrorOccured: false,
              isLoading: false
            });
          },
          //reject
          (data: any) => {
            sem.leave();
            this.setState({
              items: [],
              isLoading: false,
              isErrorOccured: true,
              errorMessage: "Something went wrong; " + data
            });
          }
        ).catch((ex) => {
          sem.leave();
          this.setState({
            items: [],
            isLoading: false,
            isErrorOccured: true,
            errorMessage: "Something went wrong; " + ex
          });
        });
    }.bind(this));
  }


  public getEvents(maxAmountOfItems: number): Promise<ICalendarEvent[]> {
    const today = new Date();
    const nextYear = new Date();
    nextYear.setHours(0, 0, 0, 0);
    nextYear.setFullYear(today.getFullYear() + 1);

    return new Promise<ICalendarEvent[]>((resolve) => {
      // Call Graph for my upcoming events
      this.client
        .api('/me/calendarview?startdatetime=' + today.toISOString() + '&enddatetime=' + nextYear.toISOString())
        .version("v1.0") // use graph
        .select("subject,webLink,onlineMeetingUrl,location,start,end,isAllDay")
        .top(maxAmountOfItems)
        .get((err, res) => {
          if (err) {
            Promise.reject(err);
          }

          // Prepare the output array
          var events: Array<ICalendarEvent> = new Array<ICalendarEvent>();

          // Map the JSON response to the output array
          res.value.map((item: any) => {
            events.push({
              Title: item.subject,
              StartDate: item.start.dateTime, // returns an object containing datetime and timezone
              EndDate: item.end.dateTime, // returns an object containing datetime and timezone
              SkypeUrl: item.onlineMeetingUrl, // returns nothing if no skype meeting
              Location: item.location, //returns an object containing displayname and location
              Url: item.webLink,
              AllDayEvent: item.isAllDay
            });
          });

          resolve(events);
        });
    });
  }

  public render(): React.ReactElement<IHelloWorldProps> {
    if (this.state.isLoading) {
      return (<div>
        <Spinner size={SpinnerSize.large} label="Loading" />
      </div>);
    }

    if (this.state.isErrorOccured) {
      return (
        <div>
          <MessageBar messageBarType={MessageBarType.warning}>
            {this.state.errorMessage}
          </MessageBar>
        </div>);
    }

    return (
      <div className={styles.helloWorld}>
        <div className={styles.container}>
          <List
            items={this.state.items}
            renderCount={this.state.items.length}
            onRenderCell={(item, index) => (
              <div>{item.Title} : {new Date(item.StartDate).toLocaleString()} - {new Date(item.EndDate).toLocaleString()}</div>
            )}
          />
        </div>
      </div>
    );
  }
}
