import { sp } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { WebPartContext } from '@microsoft/sp-webpart-base';

export class SharePointService {

  public static async registerUser(formData: any, context: WebPartContext): Promise<{ updated: boolean; userId?: number }> {
    sp.setup({
      spfxContext: {
        ...context,
        pageContext: context.pageContext,
        msGraphClientFactory: {
          getClient: async () => await context.msGraphClientFactory.getClient('3')
        }
      }
    });

    try {
      const ensureUserResult = await sp.web.ensureUser(formData.userName.loginName); 

      // Check if user already exists
      const existingUsers = await sp.web.lists.getByTitle("WellbeingCampaignRegistrationsList").items
        .filter(`Title eq '${formData.email}'`)
        .get();

      if (existingUsers.length > 0) {
        // If user exists
        const userId = existingUsers[0].Id;
        await sp.web.lists.getByTitle("WellbeingCampaignRegistrationsList").items.getById(userId).update({
          UserNameId: ensureUserResult.data.Id,
          Address: formData.address,
          Age: formData.age,
          DateOfBirth: formData.dateOfBirth,
          Country: formData.country,
          Location: formData.location,
        });
        return { updated: true, userId };
      } else {
        // If user does not exist
        await sp.web.lists.getByTitle("WellbeingCampaignRegistrationsList").items.add({
          Title: formData.email,
          UserNameId: ensureUserResult.data.Id,
          Address: formData.address,
          Age: formData.age,
          DateOfBirth: formData.dateOfBirth,
          Country: formData.country,
          Location: formData.location,
        });
        return { updated: false };
      }
    } catch (error) {
      console.error("Error in registering user:", error);
      throw new Error("Error registering user");
    }
  }

  public static async scheduleActivity(context: WebPartContext, scheduleData: any): Promise<{ updated: boolean; scheduleId?: number }> {
    sp.setup({
      spfxContext: {
        ...context,
        pageContext: context.pageContext,
        msGraphClientFactory: {
          getClient: async () => await context.msGraphClientFactory.getClient('3')
        }
      }
    });

    try {
      const existingSchedules = await sp.web.lists.getByTitle("WellbeingCampaignRegistrationsList").items
        .filter(`Title eq '${scheduleData.Title}'`)
        .get();

        console.log(existingSchedules);
        

      if (existingSchedules.length > 0) {

        const scheduleId = existingSchedules[0].Id;
        await sp.web.lists.getByTitle("WellbeingCampaignRegistrationsList").items.getById(scheduleId).update({
          ScheduleDate: scheduleData.ExerciseDate,
          StartTime: scheduleData.StartTime,
          EndTime: scheduleData.EndTime,
          ActivityName: scheduleData.ExerciseName,
        });
        return { updated: true, scheduleId };
      } else {
        // If no schedule exists
        await sp.web.lists.getByTitle("WellbeingCampaignRegistrationsList").items.add({
          // Title: scheduleData.Title, 
          // FullName: scheduleData.FullName,
          ScheduleDate: scheduleData.ExerciseDate,
          StartTime: scheduleData.StartTime,
          EndTime: scheduleData.EndTime,
          ActivityName: scheduleData.ExerciseName,
        });
        return { updated: false };
      }
    } catch (error) {
      console.error("Error in scheduling activity:", error);
      throw new Error("Error scheduling activity");
    }
  }

  public static async getActivitiesByDate(context: WebPartContext, date?: string): Promise<any[]> {
    try {
      sp.setup({
        spfxContext: {
          ...context,
          pageContext: context.pageContext,
          msGraphClientFactory: {
            getClient: async () => await context.msGraphClientFactory.getClient('3')
          }
        }
      });

      const items = await sp.web.lists.getByTitle("WellbeingCampaignRegistrationsList") 
        .items.filter(`ScheduleDate eq '${date}'`)
        .get();

      return items;
    } catch (error) {
      console.error("Error retrieving activities by date:", error);
      return [];
    }
  }

  public static async getUsers(context: WebPartContext) {
    sp.setup({
      spfxContext: {
        ...context,
        pageContext: context.pageContext,
        msGraphClientFactory: {
          getClient: async () => await context.msGraphClientFactory.getClient('3')
        }
      }
    });

    return await sp.web.lists.getByTitle("WellbeingCampaignRegistrationsList")
    .items.select("Title","UserName/Title", "Age", "Address", "Country", "Location", "DateOfBirth", "ScheduleDate", "StartTime", "EndTime", "ActivityName", "UserNameId")
      .expand("UserName")
      .get();
  }

  public static async checkUserExists(context: WebPartContext, email: string): Promise<boolean> {
    sp.setup({
      spfxContext: {
        ...context,
        pageContext: context.pageContext,
        msGraphClientFactory: {
          getClient: async () => await context.msGraphClientFactory.getClient('3')
        }
      }
    });

    const existingUsers = await sp.web.lists.getByTitle("WellbeingCampaignRegistrationsList").items
      .filter(`Title eq '${email}'`)
      .get();

    return existingUsers.length > 0;
  }

  public static async isUserInGroup(
    context: WebPartContext,
    groupName: string,
    userEmail: string
  ): Promise<boolean> {
    try {

      const group = await sp.web.siteGroups.getByName(groupName)();

      const users = await sp.web.siteGroups.getById(group.Id).users();

      return users.some((user) => user.Email === userEmail);
    } catch (error) {
      console.error("Error checking user in group:", error);
      return false;
    }
  }

  public static async isConflictActivity(
    email: string,
    exerciseDate?: string,
    startTime?: string,
    endTime?: string
  ): Promise<boolean> {
    const filter = `
      ScheduleDate eq '${exerciseDate}' and 
      ((StartTime le '${startTime}' and EndTime gt '${startTime}') or 
      (StartTime lt '${endTime}' and EndTime ge '${endTime}') or 
      (StartTime ge '${startTime}' and EndTime le '${endTime}'))
    `;
  
    const conflictingActivities = await sp.web.lists
      .getByTitle('WellbeingCampaignRegistrationsList')
      .items.filter(filter)
      .get();
  
    return conflictingActivities.length > 0;
  }

}
