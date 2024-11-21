import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IRegistrationProps {
  context : WebPartContext
  email: string;
  onRegisterSuccess: (email: string, fullName: string) => void;
}
