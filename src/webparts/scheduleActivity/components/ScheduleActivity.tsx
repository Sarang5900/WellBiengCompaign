import * as React from 'react';
import { DialogFooter, Dropdown, IPersonaProps, PrimaryButton, Stack, TextField } from '@fluentui/react';
import { IPeoplePickerContext, PeoplePicker } from '@pnp/spfx-controls-react/lib/PeoplePicker';
import { IScheduleActivityProps } from './IScheduleActivityProps';
import { SharePointService } from '../../../services/SharepointServices';
import { AnimatedDialog } from '@pnp/spfx-controls-react/lib/AnimatedDialog';
import Registration from '../../registration/components/Registration';
import RegisteredUserList from '../../registration/components/RegisteredUserList';

interface IFormErrors {
  fullName?: string;
  exerciseDate?: string;
  startTime?: string;
  endTime?: string;
  exerciseName?: string;
}

interface IActivity {
  startTime: string;
  endTime: string;
  exerciseName: string;
}

interface IScheduleActivityState {
  registrationSuccessful: boolean;
  showSuccessDialog: boolean;
  showRegistrationForm: boolean;
  email: string;
  fullName?: string;
  exerciseDate?: string;
  startTime?: string;
  endTime?: string;
  exerciseName: string;
  activities: IActivity[];
  formErrors: IFormErrors;
}

interface IScheduleActivityPropsWithSubmit extends IScheduleActivityProps {
  onActivitySubmit: (activity: IActivity) => void;
}

export default class ScheduleActivity extends React.Component<IScheduleActivityPropsWithSubmit, IScheduleActivityState> {
  public state: IScheduleActivityState = {
    registrationSuccessful: false,
    showSuccessDialog: false,
    showRegistrationForm: false,
    email: this.props.email || '',
    fullName: this.props.fullName || '',
    exerciseDate: '',
    startTime: '',
    endTime: '',
    exerciseName: '',
    activities: [],
    formErrors: {},
  };

  private handleRegistrationSuccess = async (email: string, fullName: string): Promise<void> => {
    this.setState({
      fullName: fullName,
      showRegistrationForm: false,
      email: email
    }, async () => {
      await this.fetchScheduleData();
    });
  };

  public async componentDidMount(): Promise<void> {
    await this.fetchScheduleData();
  }

  private async fetchScheduleData(): Promise<void> {
    const userExists = await SharePointService.checkUserExists(this.props.context, this.state.email);
    if (userExists) {
      const users = await SharePointService.getUsers(this.props.context);
      const userSchedule = users.find((user: { Title: string }) => user.Title === this.state.email);
      console.log(userSchedule);
      

      if (userSchedule && userSchedule.ScheduleDate) {
        const scheduleDate = new Date(userSchedule.ScheduleDate);
        if (!isNaN(scheduleDate.getTime())) {
          const formattedDate = scheduleDate.toISOString().split('T')[0];
          this.setState({
            fullName: userSchedule.UserName.Title,
            exerciseDate: formattedDate,
            startTime: userSchedule.StartTime,
            endTime: userSchedule.EndTime,
            exerciseName: userSchedule.ActivityName,
          });
        }
      }
    } else {
      this.setState({
        exerciseDate: '',
        startTime: '',
        endTime: '',
        exerciseName: '',
      });
    }
  }

  private handlePeoplePickerChange = (items: IPersonaProps[]): void => {
    if (items.length > 0) {
      const user = items[0];
      this.setState({ fullName: user.text });
    }
  };

  private handleDropdownChange = (event: React.FormEvent<HTMLDivElement>, item?: { key: string; text: string }): void => {
    if (item) {
      this.setState({ exerciseName: item.text });
    }
  };

  private validateForm = (): boolean => {
    const { exerciseDate, startTime, endTime, exerciseName} = this.state;
    const errors: IFormErrors = {};
  
    if (!exerciseDate) errors.exerciseDate = 'Exercise date is required.';
    if (!startTime) errors.startTime = 'Start time is required.';
    if (!endTime) errors.endTime = 'End time is required.';
    if (!exerciseName) errors.exerciseName = 'Exercise name is required.';

    const currentDate = new Date();
    if (exerciseDate) {
      const selectedDate = new Date(exerciseDate);
      if (selectedDate <= currentDate) {
        errors.exerciseDate = 'The exercise date must be in the future.';
      }
    }
  
    if (startTime && endTime) {
      const start = new Date(`${exerciseDate} ${startTime}`);
      const end = new Date(`${exerciseDate} ${endTime}`);
  
      if (start >= end) {
        errors.startTime = 'Start time must be before end time.';
        errors.endTime = 'End time must be after start time.';
      }
    }
  
    this.setState({ formErrors: errors });
    return Object.keys(errors).length === 0;
  };

  private confirmSubmit = async () : Promise<void> => {
    this.validateForm();

    const { fullName, email, exerciseDate, startTime, endTime, exerciseName, activities } = this.state;

    if(startTime && endTime){
      const newActivity : IActivity = {startTime, endTime, exerciseName};
      const updatedActivities = [...activities, newActivity];

      
      const hasConflict = await SharePointService.isConflictActivity(email, exerciseDate, startTime, endTime);
      if (hasConflict) {
        alert('An activity is already scheduled for the selected date and time.');
        return;
      }

      const scheduleData = {
        Title: email,
        FullName: fullName,
        ExerciseDate: exerciseDate,
        StartTime: startTime,
        EndTime: endTime,
        ExerciseName: exerciseName,
      };

      try {
        await SharePointService.scheduleActivity(this.props.context, scheduleData);

        this.setState({
          activities: updatedActivities,
          registrationSuccessful: true,
          showSuccessDialog: true,
          showRegistrationForm: false,
          email: '',
          fullName: '',
          exerciseDate: '',
          startTime: '',
          endTime: '',
          exerciseName: '',
          formErrors: {},
        });
      } catch (error) {
        console.error("Error while scheduling activity.");
      }
    }else {
      console.error("Start Time and End Time are undefined.");
    }
  };

  public handleCloseSuccessDialog = (): void => {
    this.setState({
      showSuccessDialog: false,
      registrationSuccessful: true,
    });
  };

  public render(): React.ReactElement<IScheduleActivityProps> {
    const { email, fullName, exerciseDate, startTime, endTime, exerciseName, formErrors, registrationSuccessful, showRegistrationForm } = this.state;

    const peoplePickerContext: IPeoplePickerContext = {
      absoluteUrl: this.props.context.pageContext.web.absoluteUrl,
      msGraphClientFactory: this.props.context.msGraphClientFactory,
      spHttpClient: this.props.context.spHttpClient,
    };

    return (
      <>
        {
          registrationSuccessful ? (
            <RegisteredUserList context={this.props.context} /> 
          ) : showRegistrationForm ? (
            <Registration 
              onRegisterSuccess={this.handleRegistrationSuccess} 
              email= {email}
              context={this.props.context} 
            />
          ) : (
            <Stack
              horizontalAlign="center"
              verticalAlign="center"
              styles={{
                root: {
                  height: '100vh',
                  backgroundColor: '#f7f7f7',
                },
              }}
            >
              <Stack
                tokens={{ childrenGap: 20 }}
                styles={{
                  root: {
                    maxWidth: '500px',
                    width: '90%',
                    padding: '20px',
                    boxShadow: '0px 4px 12px rgba(0, 0, 0, 0.1)',
                    borderRadius: '10px',
                    backgroundColor: '#ffffff',
                  },
                }}
              >
                <h2 style={{ textAlign: 'center', marginBottom: '20px', color: '#333' }}>Schedule Your Activity</h2>

                <Stack tokens={{ childrenGap: 15 }}>
                  <Stack horizontal tokens={{ childrenGap: 15 }}>
                    <Stack.Item styles={{ root: { width: '48%' } }}>
                      <TextField label="Email" value={email} name="email" readOnly />
                    </Stack.Item>

                    <Stack.Item styles={{ root: { width: '48%' } }}>
                      <PeoplePicker
                        context={peoplePickerContext}
                        titleText="Full Name"
                        personSelectionLimit={1}
                        onChange={this.handlePeoplePickerChange}
                        showtooltip={true}
                        required={true}
                        defaultSelectedUsers={fullName ? [fullName] : []}
                        errorMessage={formErrors.fullName}
                        disabled={true}
                      />
                    </Stack.Item>
                  </Stack>

                  <TextField
                    label="Select Date for Exercise"
                    type="date"
                    value={exerciseDate}
                    onChange={(e, newValue) => this.setState(prevState => ({
                      exerciseDate: newValue || '',
                      formErrors: { ...prevState.formErrors, exerciseDate: '' }
                    }))}
                    required={true}
                    errorMessage={formErrors.exerciseDate}
                  />

                  <Stack horizontal tokens={{ childrenGap: 15 }}>
                    <Stack.Item styles={{ root: { width: '48%' } }}>
                      <TextField
                        label="Start Time"
                        type="time"
                        value={startTime}
                        onChange={(e, newValue) => this.setState(prevState => ({
                          startTime: newValue || '',
                          formErrors: { ...prevState.formErrors, startTime: '' }
                        }))}
                        required={true}
                        errorMessage={formErrors.startTime}
                      />
                    </Stack.Item>

                    <Stack.Item styles={{ root: { width: '48%' } }}>
                      <TextField
                        label="End Time"
                        type="time"
                        value={endTime}
                        onChange={(e, newValue) => this.setState(prevState => ({
                          endTime: newValue || '',
                          formErrors: { ...prevState.formErrors, endTime: '' }
                        }))}
                        required={true}
                        errorMessage={formErrors.endTime}
                      />
                    </Stack.Item>
                  </Stack>

                  <Dropdown
                    label="Exercise Type"
                    selectedKey={exerciseName}
                    onChange={this.handleDropdownChange}
                    options={[
                      { key: 'Yoga', text: 'Yoga' },
                      { key: 'Cardio', text: 'Cardio' },
                      { key: 'Strength', text: 'Strength' },
                    ]}
                    required={true}
                    errorMessage={formErrors.exerciseName}
                  />
                </Stack>

                <DialogFooter>
                  <PrimaryButton onClick={this.confirmSubmit} text="Confirm" />
                </DialogFooter>
              </Stack>
            </Stack>
          )
        }

        {this.state.showSuccessDialog && (
          <AnimatedDialog
            isOpen={this.state.showSuccessDialog}
            onDismiss={this.handleCloseSuccessDialog}
            title="Schedule Success!"
            dialogContentProps={{
              subText: "Your activity has been successfully scheduled"
            }}
          >
            <DialogFooter>
              <PrimaryButton text="Close" onClick={this.handleCloseSuccessDialog} />
            </DialogFooter>
          </AnimatedDialog>
        )}
 
      </>
    );
  }
}
