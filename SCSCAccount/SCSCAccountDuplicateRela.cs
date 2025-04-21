using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using SalesConfigurationPlugins;
using System;
using System.ServiceModel;

namespace SCSCAccount
{
    public class SCSCAccountDuplicateRela : IPlugin
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
                // Verify the plugin is running
                tracing.Trace("Outside the context condition");
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is EntityReference)
                {
                    // Plugin is activated using Custom Action, so 'EntityReference' is used
                    EntityReference entityRef = (EntityReference)context.InputParameters["Target"];
                    if (entityRef.LogicalName != CRfF8_ScAccount.EntityLogicalName) return;
                    // We are using 'EntityReference' so we can get only logical name and GUID
                    string entityRefLogicalName = entityRef.LogicalName.ToString();
                    Guid entityRefGUID = entityRef.Id;
                    // Retrieve the full data
                    Entity original = service.Retrieve(entityRefLogicalName, entityRefGUID, new ColumnSet(true));

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
                        tracing.Trace("Value: {0}", prop.Name.ToString());

                        var value = prop.GetValue(original);

                        if (value != null)
                        {
                            prop.SetValue(clone, value);
                        }

                    }

                    clone.CRfF8_ScAccountName = "[Cloned] " + clone.CRfF8_ScAccountName;

                    service.Create(clone);

                    // Set the output parameter as "success" once completed
                    context.OutputParameters["output"] = "success";
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
