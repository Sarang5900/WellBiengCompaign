import * as React from 'react';
import { IRegistrationProps } from './IRegistrationProps';
import {
  PrimaryButton, TextField, Stack, DefaultButton, DialogFooter,
  Label,
  IPersonaProps
} from '@fluentui/react';
import { SharePointService } from '../../../services/SharepointServices';
import { AnimatedDialog } from '@pnp/spfx-controls-react/lib/AnimatedDialog';
import { IPeoplePickerContext, PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";

interface IErrors {
  userName?: string;
  email?: string;
  age?: string;
  address?: string;
  dateOfBirth?: string;
  country?: string;
  location?: string;
}


interface IRegistrationState {
  formData: {
    userName?: IPersonaProps;
    email: string;
    age: string;
    address: string;
    dateOfBirth: string;
    country: string;
    location: string;
  };
  showCustomisedAnimatedDialog: boolean;
  showSuccessDialog: boolean;
  showErrorDialog: boolean;
  showUpdateDialogue: boolean;
  registrationSuccessful: boolean;
  errors: IErrors;
}

export default class Registration extends React.Component<IRegistrationProps, IRegistrationState> {

  constructor(props: IRegistrationProps) {
    super(props);
    this.state = {
      formData: {
        userName: undefined,
        email: props.email,
        address: '',
        age: '',
        dateOfBirth: '',
        country: '',
        location: '',
      },
      showCustomisedAnimatedDialog: false,
      showSuccessDialog: false,
      showErrorDialog: false,
      showUpdateDialogue: false,  
      registrationSuccessful: false,
      errors: {},
    };
  }

  handleChange = (
    event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    newValue?: string
  ): void => {
    const { name } = event.currentTarget;
    const updatedFormData = { ...this.state.formData, [name]: newValue || '' };
    const errors = { ...this.state.errors };
  
    if (name === 'dateOfBirth') {
      const birthDate = new Date(newValue || '');
      const today = new Date();
      const age = today.getFullYear() - birthDate.getFullYear() - (today < new Date(birthDate.setFullYear(today.getFullYear())) ? 1 : 0);
  
      if (age < 18) {
        errors.dateOfBirth = "Age must be at least 18 years.";
        updatedFormData.age = '';
      } else if (birthDate > today) {
        errors.dateOfBirth = "Date of birth cannot be in the future.";
        updatedFormData.age = '';
      } else {
        errors.dateOfBirth = undefined;
        updatedFormData.age = age.toString();
      }
    }
  
    this.setState({
      formData: updatedFormData,
      errors: errors,
    });
  };
  

  handlePeoplePickerChange = (items: IPersonaProps[]): void => {
    this.setState({
      formData: {
        ...this.state.formData,
        userName: items.length > 0 ? { ...items[0] } : undefined,
      },
      errors: { 
        ...this.state.errors,
        userName: undefined,
      }
    })
  }

  validateForm = (): boolean => {
    const { formData } = this.state;
    let isValid = true;
    const errors: IErrors = {};

    if (!formData.email.trim()) {
      errors.email = "Please enter email.";
      isValid = false;
    } else if (( !/\S+@\S+\.\S+/.test(formData.email))){
      errors.email = "Please enter a valid email.";
      isValid = false;
    }

    if (!formData.userName) {
      errors.userName = "User selection is required.";
      isValid = false;
    }
    if (!formData.address.trim()) {
      errors.address = "Address is required.";
      isValid = false;
    }
    if (!formData.age.trim()) {
      errors.age = "Age is required.";
      isValid = false;
    }
    if( Number(formData.age) <=0 || Number(formData.age) >= 100   || isNaN(Number(formData.age)) ) {
      errors.age = "Please enter valid age.";
      isValid = false;
    }
    if (!formData.dateOfBirth) {
      errors.dateOfBirth = "Date of birth is required.";
      isValid = false;
    } else {
      const birthDate = new Date(formData.dateOfBirth);
      const today = new Date();

      if (birthDate > today) {
        errors.dateOfBirth = "Date of birth cannot be in the future.";
        isValid = false;
      }
    }
    if (!formData.country.trim()) {
      errors.country = "Country is required.";
      isValid = false;
    }
    if (!formData.location.trim()) {
      errors.location = "Location is required.";
      isValid = false;
    }

    this.setState({ errors });
    return isValid;
  };

  handleBlur = (event: React.FocusEvent<HTMLInputElement | HTMLTextAreaElement>): void => {
    const { name, value } = event.currentTarget;
    const errors: IErrors = { ...this.state.errors };
    const formData = { ...this.state.formData };
  
    if (name === 'email') {
      if (!value.trim()) {
        errors.email = "Email is Required.";
      } else {
        errors.email = undefined;
      }
    } else if (name === 'age') {
      if (!value.trim()) {
        errors.age = "Age is required.";
      } else {
        errors.age = undefined;
      }
    } else if (name === 'address') {
      if (!value.trim()) {
        errors.address = "Address is required.";
      } else {
        errors.address = undefined;
      }
    } else if (name === 'dateOfBirth') {
      if (!value) {
        errors.dateOfBirth = "Date of birth is required.";
      } else {
        const birthDate = new Date(value);
        const today = new Date();
        
        if (birthDate > today) {
          errors.dateOfBirth = "Date of birth cannot be in the future.";
        } else {
          const age = today.getFullYear() - birthDate.getFullYear();
          const monthDiff = today.getMonth() - birthDate.getMonth();
          const dayDiff = today.getDate() - birthDate.getDate();
  
          // Adjust age if the current month/date is before the birth month/date
          const correctedAge = (monthDiff < 0 || (monthDiff === 0 && dayDiff < 0)) ? age - 1 : age;
  
          if (correctedAge < 18) {
            errors.dateOfBirth = "Age must be 18 or older.";
            formData.age = "";
          } else {
            formData.age = correctedAge.toString();
            errors.dateOfBirth = undefined;
          }
        }
      }
    } else if (name === 'country') {
      if (!value.trim()) {
        errors.country = "Country is required.";
      } else {
        errors.country = undefined;
      }
    } else if (name === 'location') {
      if (!value.trim()) {
        errors.location = "Location is required.";
      } else {
        errors.location = undefined;
      }
    }
  
    this.setState({ errors, formData });
  };


  handleSubmit = async (event: React.MouseEvent<HTMLButtonElement>): Promise<void> => {
    event.preventDefault();

    const isFormValid = this.validateForm();
    if(!isFormValid) return;

    const { email } = this.state.formData;
    
    try {
      const userExists = await SharePointService.checkUserExists(this.props.context, email);

      if(userExists){
        this.setState({
          errors: {
            ...this.state.errors,
            email: "User already exists with the same email.",
          },
          showUpdateDialogue: true,
        });
      } else {
        this.setState({ showCustomisedAnimatedDialog: true })
      }
    } catch (error) {
      console.error("Error while checking user existence.", error);
      this.setState({
        errors: {
          ...this.state.errors,
          email: "An error occured while checking user existence.",
        },
        showErrorDialog: true,
      })
    }
  }
  
  submitForm = async (): Promise<void> => {
    try {
      await SharePointService.registerUser(this.state.formData, this.props.context);
      const { email } = this.state.formData;
      const fullName = this.state.formData.userName?.text || '';
  
      this.props.onRegisterSuccess(email, fullName);
      this.setState({
        formData: {
          userName: undefined,
          email: '',
          address: '',
          age: '',
          dateOfBirth: '',
          country: '',
          location: '',
        },
        showCustomisedAnimatedDialog: false,
        showSuccessDialog: true,
        errors: {},
      });

    } catch (error) {
      console.error("Error registering user", error);
      this.setState({
        showCustomisedAnimatedDialog: false,
        showErrorDialog: true, 
      });
    }
  };
  

  handleUpdateConfirmation = async (isConfirmed: boolean): Promise<void> => {
    if (isConfirmed) {
      await this.submitForm();
    } else {
      this.setState({ showUpdateDialogue: false });
    }
  };

  public render(): React.ReactElement<IRegistrationProps> {
    const { formData, errors, showCustomisedAnimatedDialog, showSuccessDialog, showErrorDialog, showUpdateDialogue } = this.state;

    const peoplePickerContext: IPeoplePickerContext = {
      absoluteUrl: this.props.context.pageContext.web.absoluteUrl,
      msGraphClientFactory: this.props.context.msGraphClientFactory,
      spHttpClient: this.props.context.spHttpClient,
    };

    return (
      <>
        <Stack 
          tokens={{ childrenGap: 20 }} 
          horizontalAlign="center" 
          verticalAlign="center" 
          styles={{ 
            root: { 
              height: '100vh', 
              backgroundColor: '#f7f7f7', 
            } 
          }}
          >

          <Stack
            tokens={{ childrenGap: 20 }}
            styles={{
              root: {
                maxWidth: '600px',
                width: '100%',
                padding: '20px',
                boxShadow: '0px 4px 8px rgba(0, 0, 0, 0.1)',
                borderRadius: '8px',
                backgroundColor: '#ffffff',
              },
            }}
          >
          <Label styles={{ root: { fontWeight: 'bold', fontSize: '20px', marginBottom: '20px', textAlign: 'center' } }}>
            Well-Being Registration Form
          </Label>
            <Stack horizontal={window.innerWidth > 768} tokens={{ childrenGap: 15 }}>
              <Stack.Item styles={{ root: { width: "100%", maxWidth: window.innerWidth > 768 ? "50%" : "100%", marginBottom: window.innerWidth > 768 ? 0 : '10px' } }}>
                <PeoplePicker
                  context={peoplePickerContext}
                  titleText="Full Name"
                  placeholder="Enter your Name"
                  personSelectionLimit={1}
                  showtooltip={true}
                  required={true}
                  onChange={this.handlePeoplePickerChange}
                  principalTypes={[PrincipalType.User]}
                  errorMessage={errors.userName}
                />
              </Stack.Item>
              <Stack.Item styles={{ root: { width: "100%", maxWidth: window.innerWidth > 768 ? "50%" : "100%" } }}>
                <TextField
                  label="Email"
                  name="email"
                  type="email"
                  value={formData.email}
                  onChange={this.handleChange}
                  onBlur={this.handleBlur} 
                  required
                  errorMessage={errors.email}
                  placeholder="Enter your email"
                />
              </Stack.Item>
            </Stack>

            <Stack horizontal tokens={{ childrenGap: 15 }}>
              <Stack.Item styles={{ root: { width: "50%" } }}>
                <TextField
                  label="Date of Birth"
                  name="dateOfBirth"
                  type="date"
                  value={formData.dateOfBirth}
                  onChange={this.handleChange}
                  onBlur={this.handleBlur} 
                  errorMessage={errors.dateOfBirth}
                  required
                />
              </Stack.Item>
              <Stack.Item styles={{ root: { width: "50%" } }}>
                <TextField
                  label="Age"
                  name="age"
                  type="text"
                  placeholder='Enter your age'
                  value={formData.age}
                  onChange={this.handleChange}
                  onBlur={this.handleBlur} 
                  errorMessage={errors.age}
                  readOnly
                />
              </Stack.Item>
            </Stack>

            <TextField 
              label="Address"
              multiline
              name="address"
              value={formData.address}
              onChange={this.handleChange}
              onBlur={this.handleBlur} 
              errorMessage={errors.address}
              required
              placeholder="Enter your address"
            />

            <Stack horizontal tokens={{ childrenGap: 15 }}>
              <Stack.Item styles={{ root: { width: "50%" } }}>
                <TextField
                  label="Country"
                  name="country"
                  value={formData.country}
                  onChange={this.handleChange}
                  onBlur={this.handleBlur} 
                  errorMessage={errors.country}
                  required
                  placeholder="Enter country name"
                />
              </Stack.Item>
              <Stack.Item styles={{ root: { width: "50%" } }}>
                <TextField
                  label="Location"
                  name="location"
                  value={formData.location}
                  onChange={this.handleChange}
                  onBlur={this.handleBlur} 
                  errorMessage={errors.location}
                  required
                  placeholder="Enter your location"
                />
              </Stack.Item>
            </Stack>

            <DialogFooter>
              <PrimaryButton onClick={this.handleSubmit} text="Register" /> 
            </DialogFooter>
          </Stack>

          {/* Confirmation Animated Dialog */}
          <AnimatedDialog
            hidden={!showCustomisedAnimatedDialog}
            onDismiss={() => { this.setState({ showCustomisedAnimatedDialog: false }); }}
            dialogContentProps={{
              title: 'Confirm Registration',
              subText: 'Do you want to submit your registration?',
            }}
          >
            <DialogFooter>
              <PrimaryButton onClick={this.submitForm} text="Yes" />
              <DefaultButton onClick={() => this.setState({ showCustomisedAnimatedDialog: false })} text="No" />
            </DialogFooter>
          </AnimatedDialog>

          {/* Update Confirmation Dialog */}
          <AnimatedDialog
            hidden={!showUpdateDialogue}
            onDismiss={() => { this.setState({ showUpdateDialogue: false }); }}
            dialogContentProps={{
              title: 'Update Existing Registration',
              subText: 'An existing registration was found with the same email. Do you want to update your details?',
            }}
          >
            <DialogFooter>
              <PrimaryButton onClick={() => this.handleUpdateConfirmation(true)} text="Yes" />
              <DefaultButton onClick={() => this.handleUpdateConfirmation(false)} text="No" />
            </DialogFooter>
          </AnimatedDialog>

          {/* Success and Error Dialogs */}
          <AnimatedDialog
            hidden={!showSuccessDialog}
            onDismiss={() => this.setState({ showSuccessDialog: false })}
            dialogContentProps={{ title: 'Success!', subText: 'Registration successful!' }}
          >
            <DialogFooter>
              <PrimaryButton onClick={() => this.setState({ showSuccessDialog: false })} text="Close" />
            </DialogFooter>
          </AnimatedDialog>

          <AnimatedDialog
            hidden={!showErrorDialog}
            onDismiss={() => this.setState({ showErrorDialog: false })}
            dialogContentProps={{ title: 'Error', subText: 'There was an error while submitting your form.' }}
          >
            <DialogFooter>
              <PrimaryButton onClick={() => this.setState({ showErrorDialog: false })} text="Close" />
            </DialogFooter>
          </AnimatedDialog>
        </Stack>
      </>
    );
  }
}
