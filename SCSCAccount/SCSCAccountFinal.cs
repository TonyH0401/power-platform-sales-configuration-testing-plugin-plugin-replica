using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using SalesConfigurationPlugins;
using System;
using System.ServiceModel;

namespace SCSCAccount
{
    public class SCSCAccountFinal : IPlugin
    {

        public void Execute(IServiceProvider serviceProvider)
        {
            // ============ Initialize variables used in the plugin ============
            // Contain all the metadata
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            // Functions for performing CRUD operations
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            // Tracing, logging and auditing
            ITracingService tracing = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            try
            {
                // ============ Verify plugin is running before the actual operation ============
                tracing.Trace("> Verify outside the context condition");
                // The custom action calling this plugin is a bound custom action, so to pass multiple GUID, the loop is inside the JS script, hurt performance
                // To improve performance, use unbound custom action, it pass a GUID array once, the loop is inside the plugin 
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is EntityReference)
                {
                    // Plugin is activated using Custom Action, so 'EntityReference' is used
                    EntityReference entityRef = (EntityReference)context.InputParameters["Target"];
                    if (entityRef.LogicalName != CRfF8_ScAccount.EntityLogicalName) return;
                    // We are using 'EntityReference' so we can get only logical name and GUID
                    string entityRefLogicalName = entityRef.LogicalName.ToString();
                    Guid entityRefGUID = entityRef.Id;
                    
                    // Retrieve the original full data
                    Entity original = service.Retrieve(entityRefLogicalName, entityRefGUID, new ColumnSet(true));
                    tracing.Trace("> Retrieve original data completed");

                    // Cloning process
                    var clone = new CRfF8_ScAccount();
                    var props = typeof(CRfF8_ScAccount).GetProperties();
                    foreach (var prop in props)
                    {
                        if (!prop.CanWrite ||
                            !prop.CanRead ||
                            prop.GetIndexParameters().Length > 0)
                            continue;
                        if (prop.Name.EndsWith("Id", StringComparison.OrdinalIgnoreCase) ||
                            prop.Name.StartsWith("Created", StringComparison.OrdinalIgnoreCase) ||
                            prop.Name.StartsWith("Modified", StringComparison.OrdinalIgnoreCase) ||
                            prop.Name == "OpportunityId" ||
                            prop.Name == "EntityState" ||
                            prop.Name == "StateCode" ||
                            prop.Name == "StatusCode" ||
                            prop.Name == "Attributes"
                            //prop.Name.Contains("ExtensionData") ||
                            //prop.Name.Contains("Lazy") ||
                            //prop.Name == "RowVersion" ||
                            //prop.Name == "KeyAttributes" ||
                            //prop.Name == "RelatedEntities" ||
                            //prop.Name == "FormattedValues" || 
                            //prop.Name == "LogicalName"
                            )
                            continue;
                        if (prop.Name == "CRfF8_ScAccountNumber") continue;

                        // Use this for debugging which attribute is duplicate
                        //tracing.Trace("Value: {0}", prop.Name.ToString());

                        var value = prop.GetValue(original);

                        if (value != null)
                        {
                            prop.SetValue(clone, value);
                        }
                    }
                    clone.CRfF8_ScAccountName = "[Cloned] " + clone.CRfF8_ScAccountName;
                    var clonedId = service.Create(clone);
                    tracing.Trace("> Verify cloning completed");

                    // Set the output parameter as "success" once completed
                    //context.OutputParameters["output"] = "success";

                    // Get the contacts associate with the original account
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
                    //tracing.Trace("Counter: {0}", contacts.Entities.Count.ToString());
                    tracing.Trace("> Verify getting {0} contacts completed", contacts.Entities.Count);
                    // For each contact from the original add associate it to the cloned one
                    foreach (var contact in contacts.Entities)
                    {
                        var associateRequest = new AssociateRequest
                        {
                            Target = new EntityReference("crff8_scaccount", clonedId),
                            RelatedEntities = new EntityReferenceCollection
                            {
                                new EntityReference("crff8_sccontact", contact.Id)
                            },
                            Relationship = new Relationship("crff8_SCContact_crff8_SCAccount_crff8_SCAccount")
                        };
                        service.Execute(associateRequest);

                        //var disassociateRequest = new DisassociateRequest
                        //{
                        //    Target = new EntityReference("crff8_scaccount", entityRefGUID),
                        //    RelatedEntities = new EntityReferenceCollection
                        //    {
                        //        new EntityReference("crff8_sccontact", contact.Id)
                        //    },
                        //    Relationship = new Relationship("crff8_SCContact_crff8_SCAccount_crff8_SCAccount")
                        //};
                        //service.Execute(disassociateRequest);
                    }

                    // Set the output parameter as "success" once completed
                    context.OutputParameters["output"] = "success";
                    tracing.Trace("Completed");
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
