// import * as React from "react";
// import Registration from "../../registration/components/Registration";
// import ScheduleActivity from "../../scheduleActivity/components/ScheduleActivity";
// import { Pivot, PivotItem, Stack } from "@fluentui/react";
// import { WebPartContext } from "@microsoft/sp-webpart-base";

// interface IAdminPanelProps {
//   context: WebPartContext;
// }

// export default class AdminPanel extends React.Component<IAdminPanelProps> {
//   public render(): React.ReactElement<IAdminPanelProps> {
//     const { context } = this.props;

//     return (
//       <Stack tokens={{ childrenGap: 20 }}>
//         {/* Pivot for tab-like navigation */}
//         <Pivot>
//           {/* Registration Tab */}
//           <PivotItem headerText="Register Users">
//             <Registration
//               context={context}
//               onRegisterSuccess={(email: string, fullName: string) => {
//                 alert(`User ${fullName} (${email}) has been successfully registered.`);
//               }}
//             />
//           </PivotItem>

//           {/* Schedule Activities Tab */}
//           <PivotItem headerText="Schedule Activities">
//             <ScheduleActivity
//               context={context}
//               email=""
//               fullName=""
//               onActivitySubmit={(activity) => {
//                 alert(
//                   `Activity '${activity.exerciseName}' scheduled successfully for ${activity.startTime} - ${activity.endTime}.`
//                 );
//               }}
//             />
//           </PivotItem>
//         </Pivot>
//       </Stack>
//     );
//   }
// }
