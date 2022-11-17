using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Activities;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Crm.Sdk.Messages;

namespace CustomWorkflowTasks
{
    public class UpdateCourseEndDate : CodeActivity
    {
        [Input("DateTime input")]
        [RequiredArgument]
        public InArgument<DateTime> StartDate { get; set; }

        [Input("Integer input")]
        [RequiredArgument]
        public InArgument<int> Duration { get; set; }

        [Output("DateTime output")]
        public OutArgument<DateTime> EndDate { get; set; }

        protected override void Execute(CodeActivityContext executionContext)
        {
            //Create the tracing service
            ITracingService tracingService = executionContext.GetExtension<ITracingService>();

            //Create the context
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            try
            {
                // Get the course start date (input from user)
                DateTime startDate = StartDate.Get(executionContext);

                // get the Time Zone code of user
                int? getTimeZoneCode = RetrieveCurrentUsersSettings(service);
                tracingService.Trace("User timezone code: " + getTimeZoneCode);

                // convert the UTC DateTime (startDate) into user's local DateTime format
                DateTime localDateTime = RetrieveLocalTimeFromUTCTime(startDate, getTimeZoneCode, service);
                tracingService.Trace("Course start date: " + localDateTime);

                // Get the course duration (input from user)
                int duration = Duration.Get(executionContext);
                tracingService.Trace("Course duration: " + duration);

                // Get list of business holidays already defined in D365
                List<DateTime> businessHolidays = GetBusinessHolidays(service);
                tracingService.Trace("Business holidays: " + businessHolidays);

                for (int i = 0; i < duration;)
                {
                    // update localDateTime/startDate
                    localDateTime = localDateTime.AddDays(1);

                    // check for weekend days i.e saturday,sunday and exclude it
                    if (localDateTime.DayOfWeek == DayOfWeek.Saturday || localDateTime.DayOfWeek == DayOfWeek.Sunday)
                    {
                        continue;
                    }

                    // check for business holidays and exclude it
                    foreach (var holiday in businessHolidays)
                    {
                        // check wether holiday fall on weekend. if so, then skip below if condition
                        if (!(holiday.DayOfWeek == DayOfWeek.Saturday || holiday.DayOfWeek == DayOfWeek.Sunday))
                        {
                            if (localDateTime == holiday)
                            {
                                tracingService.Trace("Business holiday: " + holiday);
                                localDateTime = localDateTime.AddDays(1);
                            }
                        }
                    }

                    i++;
                }

                // assign localDateTime/startDate to end date
                DateTime endDate = localDateTime;

                EndDate.Set(executionContext, endDate);
                tracingService.Trace("Course end date: " + endDate);
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }


        private int? RetrieveCurrentUsersSettings(IOrganizationService service)
        {
            var currentUserSettings = service.RetrieveMultiple(
                new QueryExpression("usersettings")
                {
                    ColumnSet = new ColumnSet("timezonecode"),
                    Criteria = new FilterExpression
                    {
                        Conditions =
                        {
                            new ConditionExpression("systemuserid", ConditionOperator.EqualUserId)
                        }
                    }
                }).Entities[0].ToEntity<Entity>();

            // return time zone code
            return (int?)currentUserSettings.Attributes["timezonecode"];
        }

        private DateTime RetrieveLocalTimeFromUTCTime(DateTime utcTime, int? timeZoneCode, IOrganizationService service)
        {
            if (!timeZoneCode.HasValue)
            {
                return DateTime.Now;
            }

            var request = new LocalTimeFromUtcTimeRequest
            {
                TimeZoneCode = timeZoneCode.Value,
                UtcTime = utcTime.ToUniversalTime()
            };

            var response = (LocalTimeFromUtcTimeResponse)service.Execute(request);

            // return local time
            return response.LocalTime;
        }

        private List<DateTime> GetBusinessHolidays(IOrganizationService service)
        {
            List<DateTime> holidays = new List<DateTime>();

            QueryExpression query = new QueryExpression("calendar");
            query.Criteria.AddCondition("name", ConditionOperator.Equal, "Business Holidays");
            EntityCollection collection = service.RetrieveMultiple(query);

            if (collection.Entities.Count != 1)
            {
                return holidays;
            }

            Entity calendar = collection.Entities[0];

            if (!calendar.Contains("calendarrules"))
            {
                return holidays;
            }

            EntityCollection rules = calendar.GetAttributeValue<EntityCollection>("calendarrules");

            foreach (Entity rule in rules.Entities)
            {
                holidays.Add(rule.GetAttributeValue<DateTime>("starttime"));
            }

            return holidays;
        }


        //private DateTime RetrieveUTCTimeFromLocalTime(DateTime localTime, int? timeZoneCode, IOrganizationService service)
        //{
        //    if (!timeZoneCode.HasValue)
        //    {
        //        return DateTime.Now;
        //    }

        //    var request = new UtcTimeFromLocalTimeRequest
        //    {
        //        TimeZoneCode = timeZoneCode.Value,
        //        LocalTime = localTime
        //    };

        //    var response = (UtcTimeFromLocalTimeResponse)service.Execute(request);

        //    // return utc time
        //    return response.UtcTime;
        //}
    }
}
