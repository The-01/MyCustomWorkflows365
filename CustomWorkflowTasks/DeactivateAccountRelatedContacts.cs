using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;

using System;
using System.Activities;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;

namespace CustomWorkflowTasks
{
    public class DeactivateAccountRelatedContacts : CodeActivity
    {
        #region

        //[Input("OptionSetValue input")]
        //[RequiredArgument]
        //[AttributeTarget("account", "statecode")]
        //public InArgument<OptionSetValue> StateCode { get; set; }

        [Input("OptionSetValue input")]
        [RequiredArgument]
        [AttributeTarget("account", "statuscode")]
        public InArgument<OptionSetValue> StatusCode { get; set; }

        #endregion

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
                // Get the current account id
                Guid accountId = context.PrimaryEntityId;

                // query to retrieve all related contacts to current account
                var fetchXml = 
                    $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                        <entity name='contact'>
                            <attribute name='fullname' />
                            <attribute name='telephone1' />
                            <attribute name='contactid' />
                            <order attribute='fullname' descending='false' />
                            <filter>
                                <condition attribute='parentcustomerid' operator='eq' value='{accountId}' uitype='account' />
                            </filter>
                        </entity>
                    </fetch>";


                ////Get the state code i.e active or deactive(0 or 1)
                //int stateCode = StateCode.Get<OptionSetValue>(executionContext).Value;
                //tracingService.Trace("State code: " + stateCode);

                //Get the status code i.e active or deactive(1 or 2)
                int statusCode = StatusCode.Get<OptionSetValue>(executionContext).Value;
                tracingService.Trace("Status code: " + statusCode);

                if (statusCode == 1)
                {
                    ActivateRecord(fetchXml, service, tracingService);
                }
                else
                {
                    DeactivateRecord(fetchXml, service, tracingService);
                }
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                throw new Exception(ex.Message);
            }
        }


        //Deactivate a record
        public static void DeactivateRecord(string fetchXml, IOrganizationService service, ITracingService trace)
        {
            // Retrieve all account related contacts
            EntityCollection collection = service.RetrieveMultiple(new FetchExpression(fetchXml));

            trace.Trace("Inside deactivate method...");
            foreach (var entity in collection.Entities)
            {
                trace.Trace(entity.Attributes["fullname"].ToString());
                SetStateRequest setStateRequest = new SetStateRequest()
                {
                    EntityMoniker = new EntityReference
                    {
                        Id = entity.Id,
                        LogicalName = entity.LogicalName,
                    },
                    State = new OptionSetValue(1),
                    Status = new OptionSetValue(2)
                };

                service.Execute(setStateRequest);
            }
        }

        //Activate a record
        public static void ActivateRecord(string fetchXml, IOrganizationService service, ITracingService trace)
        {
            // Retrieve all account related contacts
            EntityCollection collection = service.RetrieveMultiple(new FetchExpression(fetchXml));

            trace.Trace("Inside activate method...");
            foreach (var entity in collection.Entities)
            {
                trace.Trace(entity.Attributes["fullname"].ToString());
                SetStateRequest setStateRequest = new SetStateRequest()
                {
                    EntityMoniker = new EntityReference
                    {
                        Id = entity.Id,
                        LogicalName = entity.LogicalName,
                    },
                    State = new OptionSetValue(0),
                    Status = new OptionSetValue(1)
                };

                service.Execute(setStateRequest);
            }
        }
    }
}
