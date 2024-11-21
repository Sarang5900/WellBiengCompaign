import * as React from "react";
import { DetailsList, DetailsListLayoutMode, IColumn } from "@fluentui/react";
import { SharePointService } from "../../../services/SharepointServices";
import { WebPartContext } from "@microsoft/sp-webpart-base";

interface IRegisteredUsersListProps {
  context: WebPartContext;
}

interface IRegisteredUsersListState {
  items: IUserItem[];
}

interface IUserItem {
  key: number;
  email: string;
  fullName: string;
  age: number;
  address: string;
  country: string;
  location: string;
  dateOfBirth: string;
  activityDate: string;
  startTime: string;
  endTime: string;
  activityName: string;
}

interface IUser {
  Title: string; 
  UserName?: { Title: string }
  Age: number;
  Address: string;
  Country: string;
  Location: string;
  DateOfBirth: string;
  ScheduleDate: string;
  StartTime: string;
  EndTime: string;
  ActivityName: string;
}

const formatDate = (dateString: string): string => {
  const date = new Date(dateString);
  const day = date.getDate().toString().padStart(2, "0");
  const month = (date.getMonth() + 1).toString().padStart(2, "0"); // Months are zero-based
  const year = date.getFullYear();
  return `${day}-${month}-${year}`;
};

class RegisteredUsersList extends React.Component<
  IRegisteredUsersListProps,
  IRegisteredUsersListState
> {
  private columns: IColumn[];

  constructor(props: IRegisteredUsersListProps) {
    super(props);

    this.state = {
      items: [],
    };

    this.columns = [
      {
        key: "column1",
        name: "Email",
        fieldName: "email",
        minWidth: 100,
        maxWidth: 150,
        isResizable: true,
        onColumnClick: this.handleColumnClick,
      },
      {
        key: "column2",
        name: "Full Name",
        fieldName: "fullName",
        minWidth: 50,
        maxWidth: 100,
        isResizable: true,
        onColumnClick: this.handleColumnClick,
      },
      {
        key: "column3",
        name: "Date Of Birth",
        fieldName: "dateOfBirth",
        minWidth: 50,
        maxWidth: 100,
        isResizable: true,
        onColumnClick: this.handleColumnClick,
      },
      {
        key: "column4",
        name: "Age",
        fieldName: "age",
        minWidth: 50,
        maxWidth: 50,
        isResizable: true,
        onColumnClick: this.handleColumnClick,
      },
      {
        key: "column5",
        name: "Address",
        fieldName: "address",
        minWidth: 150,
        maxWidth: 200,
        isResizable: true,

        onColumnClick: this.handleColumnClick,
      },
      {
        key: "column6",
        name: "Country",
        fieldName: "country",
        minWidth: 50,
        maxWidth: 100,
        isResizable: true,
        onColumnClick: this.handleColumnClick,
      },
      {
        key: "column7",
        name: "Location",
        fieldName: "location",
        minWidth: 50,
        maxWidth: 100,
        isResizable: true,
        onColumnClick: this.handleColumnClick,
      },
      {
        key: "column8",
        name: "Activity Date",
        fieldName: "activityDate",
        minWidth: 50,
        maxWidth: 100,
        isResizable: true,
        onColumnClick: this.handleColumnClick,
      },
      {
        key: "column9",
        name: "Start Time",
        fieldName: "startTime",
        minWidth: 50,
        maxWidth: 100,
        isResizable: true,
        onColumnClick: this.handleColumnClick,
      },
      {
        key: "column10",
        name: "End Time",
        fieldName: "endTime",
        minWidth: 50,
        maxWidth: 100,
        isResizable: true,
        onColumnClick: this.handleColumnClick,
      },
      {
        key: "column11",
        name: "Activity Name",
        fieldName: "activityName",
        minWidth: 50,
        maxWidth: 100,
        isResizable: true,
        onColumnClick: this.handleColumnClick,
      },
    ];
  }

  private handleColumnClick = (
    event: React.MouseEvent<HTMLElement>,
    column: IColumn
  ): void => {
    const { items } = this.state;

    // Ensure the fieldName is a key of IUserItem
    const fieldName = column.fieldName as keyof IUserItem;
    if (!fieldName) return;

    const isSortedDescending = column.isSortedDescending;
    const newSortedDescending = !isSortedDescending;

    const sortedItems = items.slice().sort((a, b) => {
      const fieldA = a[fieldName];
      const fieldB = b[fieldName];

      if (fieldA < fieldB) {
        return newSortedDescending ? 1 : -1;
      }
      if (fieldA > fieldB) {
        return newSortedDescending ? -1 : 1;
      }
      return 0;
    });

    // Update the column with sorting state
    this.columns = this.columns.map((col) => {
      col.isSorted = col.key === column.key;
      col.isSortedDescending =
        col.key === column.key ? newSortedDescending : undefined;
      return col;
    });

    this.setState({ items: sortedItems });
  };

  public async componentDidMount(): Promise<void> {
    const users = await SharePointService.getUsers(this.props.context);

    const items = users.map((user: IUser, index: number) => ({
      key: index,
      email: user.Title,
      fullName: user.UserName ? user.UserName.Title : "",
      age: user.Age,
      address: user.Address,
      country: user.Country,
      location: user.Location,
      dateOfBirth: formatDate(user.DateOfBirth),
      activityDate: formatDate(user.ScheduleDate),
      startTime: user.StartTime,
      endTime: user.EndTime,
      activityName: user.ActivityName,
    }));

    this.setState({ items });
  }

  public render(): JSX.Element {
    const { items } = this.state;

    return (
      <div>
        <h2>Registered Users</h2>
        <DetailsList
          items={items}
          columns={this.columns}
          setKey="set"
          layoutMode={DetailsListLayoutMode.fixedColumns}
          selectionPreservedOnEmptyClick={true}
          ariaLabelForSelectionColumn="Toggle selection"
          ariaLabelForSelectAllCheckbox="Toggle selection for all items"
          checkButtonAriaLabel="Row checkbox"
        />
      </div>
    );
  }
}

export default RegisteredUsersList;
