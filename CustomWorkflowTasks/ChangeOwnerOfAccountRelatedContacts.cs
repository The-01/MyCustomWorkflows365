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
    public class ChangeOwnerOfAccountRelatedContacts : CodeActivity
    {
        #region

        [Input("Lookup input")]
        [RequiredArgument]
        [ReferenceTarget("systemuser")]
        public InArgument<EntityReference> AccountOwner { get; set; }

        #endregion

        protected override void Execute(CodeActivityContext executionContext)
        {
            // Create the tracing service
            ITracingService tracingService = executionContext.GetExtension<ITracingService>();

            // Create the context
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
                            <attribute name='contactid' />
                            <order attribute='fullname' descending='false' />
                            <filter>
                                <condition attribute='parentcustomerid' operator='eq' value='{accountId}' uitype='account' />
                            </filter>
                        </entity>
                    </fetch>";


                // Get the account owner id after account is assigned to another owner
                Guid accountOwnerId = AccountOwner.Get<EntityReference>(executionContext).Id;
                tracingService.Trace("ownerid: " + accountOwnerId);

                SetContactOwner(accountOwnerId, fetchXml, service, tracingService);
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                throw new Exception(ex.Message);
            }
        }


        // Change the owner of all account related contacts
        public static void SetContactOwner(Guid accountOwnerId, string fetchXml, IOrganizationService service, ITracingService tracingService)
        {
            // Retrieve all account related contacts
            EntityCollection collection = service.RetrieveMultiple(new FetchExpression(fetchXml));

            tracingService.Trace("Inside SetContactOwner method...");
            foreach (var entity in collection.Entities)
            {
                Entity contact = new Entity(entity.LogicalName);
                contact.Id = entity.Id;
                contact["ownerid"] = accountOwnerId;

                service.Update(contact);
            }
        }
    }
}
