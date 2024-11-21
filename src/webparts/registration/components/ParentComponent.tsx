import * as React from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { Stack, TextField, PrimaryButton, Pivot, PivotItem, Label, IBreadcrumbItem, Breadcrumb, Link } from "@fluentui/react";
import Registration from "./Registration";
import ScheduleActivity from "../../scheduleActivity/components/ScheduleActivity";
import RegisteredUsersList from "./RegisteredUserList";
import { SharePointService } from "../../../services/SharepointServices";

interface IUser {
  Title: string;  // assuming 'Title' is the email field
  FullName: string;
}

interface IParentState {
  isAdmin: boolean | undefined;
  email: string;
  fullName: string;
  registrationSuccessful: boolean;
  userExists: boolean;
  isLoading: boolean;
  showRegistration: boolean;
  errors: {email?: string};
  breadcrumbItems: IBreadcrumbItem[];
}

interface IParentComponentProps {
  context: WebPartContext;
}

class ParentComponent extends React.Component<IParentComponentProps, IParentState> {

  private handleBreadcrumbClick: (ev?: React.MouseEvent<HTMLElement>, item?: IBreadcrumbItem) => void = (
    ev?: React.MouseEvent<HTMLElement>,
    item?: IBreadcrumbItem
  ) => {
    if (item?.key === 'home') {

      this.setState({
        isAdmin: undefined,
        email: "",
        fullName: "",
        registrationSuccessful: false,
        userExists: false,
        isLoading: false,
        showRegistration: false,
        errors: {},
        breadcrumbItems: [{ text: 'Go To Home', key: 'home', onClick: this.handleBreadcrumbClick }],
      });
    }
  };

  public state: IParentState = {
    isAdmin: undefined,
    email: "",
    fullName: "",
    registrationSuccessful: false,
    userExists: false,
    isLoading: false,
    showRegistration: false,
    errors: {},
    breadcrumbItems: [{ text: 'Go To Home', key: 'home', onClick: this.handleBreadcrumbClick }],
  };

  public async componentDidMount(): Promise<void> {
    this.setState({ email: "", fullName: "" });
  }

  private handleEmailChange = (
    event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    newValue?: string
  ): void => {
    const email = newValue || "";
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/; 
    const errors = { ...this.state.errors };
  
    if (!email) {
      errors.email = "Email is required.";
    } else if (!emailRegex.test(email)) {
      errors.email = "Please enter a valid email address.";
    } else {
      errors.email = undefined; 
    }
  
    this.setState({ email, errors });
  };

  private checkUser = async (): Promise<void> => {
    const email = this.state.email;
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    const errors = { ...this.state.errors };

    if (!email) {
      errors.email = "Email is required.";
    } else if (!emailRegex.test(email)) {
      errors.email = "Please enter a valid email address.";
    } else {
      errors.email = undefined;
    }

    if (errors.email) {
      this.setState({ errors });
      return;
    }

    this.setState({ isLoading: true, showRegistration: false });

    const userExists = await SharePointService.checkUserExists(this.props.context, email);

    if (userExists) {
      const users = await SharePointService.getUsers(this.props.context);
      const user = users.find((userItem: IUser) => userItem.Title === email);
      if (user) {
        const isAdmin = this.checkIfAdmin(email);
        this.setState({
          isAdmin: isAdmin,  
          registrationSuccessful: true,
          userExists: true,
          fullName: user.FullName || "",
          isLoading: false,
        });
      }
    } else {
      this.setState({
        userExists: false,
        showRegistration: true,
        isLoading: false,
      });
    }
  };

  private checkIfAdmin(email: string): boolean {
    const adminEmails = ["sarangraut5900@gmail.com", "amruta123@gmail.com"];
    return adminEmails.includes(email);
  }

  private handleRegistrationSuccess = (email: string, fullName: string): void => {
    const isAdmin = this.checkIfAdmin(email);
    this.setState({
      registrationSuccessful: true,
      email,
      fullName,
      showRegistration: false, 
      userExists: true, 
      isAdmin,
    });
  };

  
  private handleActivitySubmit = (): void => {
    this.setState({ registrationSuccessful: true });
  };

  public render(): React.ReactNode {
    const { isAdmin, email, fullName, userExists, isLoading, showRegistration, registrationSuccessful, errors, breadcrumbItems } = this.state;
  
    if (isLoading) {
      return <div>Loading...</div>;
    }

    const breadcrumb = (
      <Breadcrumb
        items={breadcrumbItems}
        onRenderItem={(items, defaultRender) => (
          <Link
            key={items?.key}
            onClick={items?.onClick}
            style={{
              padding: "1px 6px",
              border: "1px solid transparent",
              borderRadius: "5px",
              textDecoration: "none",
              color: "#0078d4", 
              fontWeight: "600",
              transition: "all 0.3s ease", 
            }}
            onMouseEnter={(e) => {
              e.currentTarget.style.border = "1px solid #0078d4"; // blue border on hover
              e.currentTarget.style.backgroundColor = "#f3f2f1"; 
            }}
            onMouseLeave={(e) => {
              e.currentTarget.style.border = "1px solid transparent"; 
              e.currentTarget.style.backgroundColor = "transparent"; 
            }}
          >
            {defaultRender ? defaultRender(items) : items?.text}
          </Link>
        )}
      />
    );
    
  
    // If user doesn't exist and is not yet registered, show email input and registration form
    if (!userExists && !registrationSuccessful) {
      return (
        <Stack tokens={{ childrenGap: 20 }} styles={{ root: { padding: "20px" } }}>
          {!showRegistration && (
            <Stack 
              tokens={{ childrenGap: 20 }} 
              styles={{ root: { padding: "20px", maxWidth: "500px", margin: "auto", textAlign: "center" } }}
            >
              <Label styles={{ root: { color: "#0078d4", fontWeight: "bold", fontSize: "24px" } }}>
                Welcome!
              </Label>
              <Label styles={{ root: { color: "#333", fontSize: "18px" } }}>
                Enter your email to get started!
              </Label>
              <TextField
                label="Email Address:"
                value={email}
                onChange={this.handleEmailChange}
                placeholder="example@domain.com"
                underlined
                styles={{ root: { marginTop: "10px" } }}
                errorMessage={errors?.email}
              />
              <PrimaryButton 
                text="Submit" 
                onClick={this.checkUser} 
                styles={{ root: { backgroundColor: "#0078d4", borderColor: "#005a9e", color: "#fff" } }}
              />
            </Stack>
          )}
          {showRegistration && (
            <>
              {breadcrumb}
              <Registration
                context={this.props.context}
                email={email}
                onRegisterSuccess={this.handleRegistrationSuccess}
              />
            </>
          )}
        </Stack>
      );
    }
  
    // If user registration is successful, check admin status and render accordingly
    if (registrationSuccessful) {
      if (isAdmin === true) {
        return (
          <Stack tokens={{ childrenGap: 20 }} styles={{ root: { padding: "20px" } }}>
            {breadcrumb}
            <Pivot>
              <PivotItem headerText="Register Users" itemIcon="Contact">
                <Registration
                  context={this.props.context}
                  email={""}
                  onRegisterSuccess={this.handleRegistrationSuccess}
                />
              </PivotItem>
  
              <PivotItem headerText="Schedule Activities" itemIcon="Calendar">
                <ScheduleActivity
                  email={email}
                  fullName={fullName}
                  context={this.props.context}
                  onActivitySubmit={this.handleActivitySubmit}
                />
              </PivotItem>
  
              <PivotItem headerText="View Registered Users" itemIcon="View">
                <RegisteredUsersList context={this.props.context} />
              </PivotItem>
            </Pivot>
          </Stack>
        );
      }
  
      // If user is not an admin, directly show ScheduleActivity
      if (isAdmin === false) {
        return (
          <>
            {breadcrumb}
            <ScheduleActivity
              email={email}
              fullName={fullName}
              context={this.props.context}
              onActivitySubmit={this.handleActivitySubmit}
            />
          </>
        );
      }
    }
  
    return undefined;
  }
}

export default ParentComponent;
