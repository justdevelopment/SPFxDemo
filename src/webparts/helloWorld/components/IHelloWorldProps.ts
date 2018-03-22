import { ICalendarEvent } from './HelloWorld';
import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IHelloWorldProps {
  context: WebPartContext;
}

export interface IHelloWorldState {
  items: ICalendarEvent[];
  isLoading?: boolean;
  isErrorOccured?: boolean;
  errorMessage?: string;
}
