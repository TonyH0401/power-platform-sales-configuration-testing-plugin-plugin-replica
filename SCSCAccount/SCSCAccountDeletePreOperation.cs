using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using SalesConfigurationPlugins;
using System;
using System.ServiceModel;

namespace SCSCAccount
{
    public class SCSCAccountDeletePreOperation : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            // Contain all the meta data
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            // Perform CRUD operations
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            // Logging
            ITracingService tracing = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            try
            {
                // The delete process uses 'EntityReference'
                // Verify the plugin is running
                tracing.Trace("Pre delete condition");
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is EntityReference)
                {
                    // Plugin is activated using Custom Action, so 'EntityReference' is used
                    EntityReference entity = (EntityReference)context.InputParameters["Target"];
                    if (entity.LogicalName != CRfF8_ScAccount.EntityLogicalName) return;
                    // We are using 'EntityReference' so we can get only logical name and GUID
                    string entityRefLogicalName = entity.LogicalName.ToString();
                    Guid entityRefGUID = entity.Id;
                    tracing.Trace("Entity GUID: {0}", entityRefGUID);
                    //// Retrieve the full data
                    //Entity original = service.Retrieve(entityRefLogicalName, entityRefGUID, new ColumnSet(true));

                    // Get the contacts associate with the original account, use the relationship table name
                    var fetchXml = $@"
                            <fetch>
                              <entity name='crff8_sccontact'>
                                <link-entity name='crff8_sccontact_crff8_scaccount' from='crff8_sccontactid' to='crff8_sccontactid' intersect='true'>
                                  <filter>
                                    <condition attribute='crff8_scaccountid' operator='eq' value='{entityRefGUID}' />
                                  </filter>
                                </link-entity>
                              </entity>
                            </fetch>";
                    var contacts = service.RetrieveMultiple(new FetchExpression(fetchXml));
                    tracing.Trace("Verify Associate Counter: {0}", contacts.Entities.Count.ToString());

                    if (contacts.Entities.Count != 0)
                    {
                        // For each contact add the disassociate to the clone one, use the relationship name
                        foreach (var contact in contacts.Entities)
                        {
                            var disassociateRequest = new DisassociateRequest
                            {
                                Target = new EntityReference("crff8_scaccount", entityRefGUID),
                                RelatedEntities = new EntityReferenceCollection
                            {
                                new EntityReference("crff8_sccontact", contact.Id)
                            },
                                Relationship = new Relationship("crff8_SCContact_crff8_SCAccount_crff8_SCAccount")
                            };
                            service.Execute(disassociateRequest);
                        }
                    }

                    // Set the output parameter as "success" once completed
                    context.OutputParameters["output"] = "success";
                    tracing.Trace(context.OutputParameters["output"].ToString());
                }
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                // This is for `Retrieve()` is null case
                if (ex.Detail.ErrorCode == -2147220969)
                {
                    tracing.Trace("Exception: the Retrieve() method returns null.");
                    throw new InvalidPluginExecutionException("Exception: the Retrieve() method returns null.");
                }
                // Others
                tracing.Trace("FaultException Code: {0}", ex.Detail.ErrorCode.ToString());
                tracing.Trace("FaultException Message: {0}", ex.Message.ToString());
                throw new InvalidPluginExecutionException("There is FaultException.", ex);
            }
            catch (Exception ex)
            {
                tracing.Trace("Exception Message: {0}", ex.Message.ToString());
                throw;
            }
        }
    }
}
