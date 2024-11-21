import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IScheduleActivityProps {
  context : WebPartContext;
  email: string;
  fullName: string;
}
