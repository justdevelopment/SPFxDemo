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

  constructor(props: IHelloWorldProps) {
    super(props);

    this.state = {
      items: []
    };
  }


  public componentDidMount() {
    this.getEvents(10).then(
      //resolve
      (apps: ICalendarEvent[]) => {
        this.setState({
          items: apps,
          isErrorOccured: false,
          isLoading: false
        });
      },
      //reject
      (data: any) => {
        console.log(data);
        this.setState({
          items: [],
          isLoading: false,
          isErrorOccured: true,
          errorMessage: "Something went wrong; " + data
        });
      }
    ).catch((ex) => {
      console.log(ex);
      this.setState({
        items: [],
        isLoading: false,
        isErrorOccured: true,
        errorMessage: "Something went wrong; " + ex
      });
    });
  }

  public getEvents(maxAmountOfItems: number): Promise<ICalendarEvent[]> {
    // get graph client from context, this was added in SPFx 1.4.1
    const client: MSGraphClient = this.props.context.serviceScope.consume(MSGraphClient.serviceKey);
    const today = new Date();
    const nextYear = new Date();
    nextYear.setHours(0, 0, 0, 0);
    nextYear.setFullYear(today.getFullYear() + 1);

    return new Promise<ICalendarEvent[]>((resolve) => {
      // Call Graph for my upcoming events
      client
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
