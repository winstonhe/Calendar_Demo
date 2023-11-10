using Microsoft.Xrm.Tooling.Connector;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Crm.Sdk.Messages;
using System;
using System.Collections.Generic;

class Program
{
  
    static void Main(string[] args)
    {
        // Connection string to your Dynamics 365 instance
        string connectionString = "AuthType=Office365;Url=https://winston.crm5.dynamics.com;Username=renhe@####.onmicrosoft.com;Password=######!";

        // Establish connection to Dynamics 365
        CrmServiceClient organizationProxy = new CrmServiceClient(connectionString);
        if (organizationProxy.IsReady)
        {
            // WhoAmI Request
            WhoAmIRequest whoAmIRequest = new WhoAmIRequest();
            WhoAmIResponse whoAmIResponse = (WhoAmIResponse)organizationProxy.Execute(whoAmIRequest);

            // Display the response
            Console.WriteLine("Logged in user ID: " + whoAmIResponse.UserId);

            Guid userid = ((WhoAmIResponse)organizationProxy.Execute(new WhoAmIRequest())).UserId;

            //clear calendar
            ClearCalenderRules(organizationProxy, userid);

            // add calendar
            AddCalenderRules(organizationProxy, userid);

            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }
    }

    public static void AddCalenderRules(IOrganizationService organizationService, Guid userid)
    {
        if (userid != Guid.Empty)
        {

            Entity systemUserEntity = organizationService.Retrieve("systemuser", userid, new ColumnSet(new String[] { "calendarid" }));
            Entity userCalendarEntity = organizationService.Retrieve("calendar", ((Microsoft.Xrm.Sdk.EntityReference)(systemUserEntity.Attributes["calendarid"])).Id, new ColumnSet(true));
            EntityCollection calendarRules = (EntityCollection)userCalendarEntity.Attributes["calendarrules"];

            Entity newInnerCalendar = new Entity("calendar");
            newInnerCalendar.Attributes["businessunitid"] = new EntityReference("businessunit", ((Microsoft.Xrm.Sdk.EntityReference)(userCalendarEntity["businessunitid"])).Id);
            Guid innerCalendarId = organizationService.Create(newInnerCalendar);

            //Create a new calendar rule and assign the inner calendar id to it
            Entity calendarRule = new Entity("calendarrule");
            calendarRule.Attributes["duration"] = 1440;
            calendarRule.Attributes["extentcode"] = 1;
            //Create a pattern of Mon-Fri
            calendarRule.Attributes["pattern"] = "FREQ=DAILY;COUNT=1";
            calendarRule.Attributes["rank"] = 0;
            //85 = UK Time Code
            calendarRule.Attributes["timezonecode"] = 85;
            calendarRule.Attributes["innercalendarid"] = new EntityReference("calendar", innerCalendarId);

            DateTime today = DateTime.UtcNow;
            calendarRule.Attributes["starttime"] = new DateTime(today.Year, today.Month, today.Day, 0, 0, 0, DateTimeKind.Utc);
            calendarRules.Entities.Add(calendarRule);

            //Assign all the calendar rule back to the user calendar
            userCalendarEntity.Attributes["calendarrules"] = calendarRules;
            //Please refer to here for Calander Types https://msdn.microsoft.com/en-us/library/dn689038.aspx
            userCalendarEntity.Attributes["type"] = new OptionSetValue(-1);
            organizationService.Update(userCalendarEntity);

            //Creates a new Calendar Rule of working hour
            Entity calendarRule_working = new Entity("calendarrule");
            
            calendarRule_working.Attributes["duration"] = 120;
            calendarRule_working.Attributes["effort"] = 1.0;
            calendarRule_working.Attributes["issimple"] = true;
           
            calendarRule_working.Attributes["offset"] = 60;
            calendarRule_working.Attributes["rank"] = 0;
            calendarRule_working.Attributes["subcode"] = 1;
            calendarRule_working.Attributes["timecode"] = 0;
            calendarRule_working.Attributes["timezonecode"] = 210;
            calendarRule_working.Attributes["calendarid"] = new EntityReference("calendar", innerCalendarId);


            EntityCollection innerCalendarRules = new EntityCollection();
            innerCalendarRules.EntityName = "calendarrule";
            innerCalendarRules.Entities.Add(calendarRule_working);

            newInnerCalendar.Attributes["calendarrules"] = innerCalendarRules;
            newInnerCalendar.Attributes["calendarid"] = innerCalendarId;
            organizationService.Update(newInnerCalendar);

            //Creates a new Calendar Rule of  break
            Entity calendarRule_break = new Entity("calendarrule");
      
            calendarRule_break.Attributes["duration"] = 15;
            calendarRule_break.Attributes["effort"] = 1.0;
            calendarRule_break.Attributes["issimple"] = true;
          
            calendarRule_break.Attributes["offset"] = 180;
            calendarRule_break.Attributes["rank"] = 0;
            calendarRule_break.Attributes["subcode"] = 4;
            calendarRule_break.Attributes["timecode"] = 2;
            calendarRule_break.Attributes["timezonecode"] = 210;
            calendarRule_break.Attributes["calendarid"] = new EntityReference("calendar", innerCalendarId);

            EntityCollection innerCalendarRules_break = new EntityCollection();            
            innerCalendarRules.Entities.Add(calendarRule_break);

            newInnerCalendar.Attributes["calendarrules"] = innerCalendarRules;
            newInnerCalendar.Attributes["calendarid"] = innerCalendarId;
            organizationService.Update(newInnerCalendar);

            Console.WriteLine("Work Hours Added to " + userid);

        }
    }

    public static void ClearCalenderRules(IOrganizationService organizationService, Guid userId)
    {
        if (userId != null)
        {
            //Retrieves all CalanderId's for all the users
            string fetchxml = "<?xml version='1.0'?>" +
                            "<fetch distinct='true' mapping='logical' output-format='xml-platform' version='1.0'>" +
                            "<entity name='calendar' >" +
                            "    <attribute name='calendarid' />" +
                            "    <filter type='and' >" +
                            "      <condition attribute='primaryuserid' operator='eq' value='" + userId + "' />" +
                            "    </filter>" +
                            "  </entity>" +
                            "</fetch>";

            EntityCollection result = organizationService.RetrieveMultiple(new FetchExpression(fetchxml));

            Console.WriteLine("There are {0} entities found", result.Entities.Count);

            foreach (var c in result.Entities)
            {
                Entity Calendar = organizationService.Retrieve("calendar", ((Guid)c.Attributes["calendarid"]), new ColumnSet(true));
                EntityCollection calendarRules = (EntityCollection)Calendar.Attributes["calendarrules"];

                int num = 0;
                List<int> list = new List<int>();

                foreach (Entity current in calendarRules.Entities)
                {
                    list.Add(num);
                    num++;
                }

                list.Sort();
                list.Reverse();

                for (int i = 0; i < list.Count; i++)
                {
                    //Remove all Calander Rules from Collection
                    calendarRules.Entities.Remove(calendarRules.Entities[list[i]]);
                }

                //Assign Calander Rules to empty Entity Collection
                Calendar.Attributes["calendarrules"] = calendarRules;
                organizationService.Update(Calendar);

                Console.WriteLine("Work Hours Deleted for " + userId);
            }
        }
    }
}


