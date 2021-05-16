using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;

namespace BasicPlugin
{
    public class CreateLaborForPhone : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            // Obtain the tracing service
            ITracingService tracingService =
            (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            // Obtain the execution context from the service provider.  
            IPluginExecutionContext context = (IPluginExecutionContext)
                serviceProvider.GetService(typeof(IPluginExecutionContext));

            // The InputParameters collection contains all the data passed in the message request.  
            if (context.InputParameters.Contains("Target") &&
                context.InputParameters["Target"] is Entity)
            {
                // Obtain the target entity from the input parameters.  
                Entity entity = (Entity)context.InputParameters["Target"];

                // Obtain the organization service reference which you will need for  
                // web service calls.  
                IOrganizationServiceFactory serviceFactory =
                    (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

                try
                {
                    // Create a task activity to follow up with the account customer in 7 days. 
                    Entity labor = new Entity("ritsc_caselabor");

                    labor["ritsc_TimeSpent"] = context.OutputParameters["actualdurationminutes"];
                    labor["ritsc_comment"] = context.OutputParameters["description"];
                    labor["ritsc_Category"] = "Communication";
                    labor["ritsc_name"] = "Temp";
                    labor["ritsc_Billable"] = "Billable";

                    // Refer to the account in the task activity.
                    if (context.OutputParameters.Contains("regardingobjectid"))
                    {
                        Guid regardingobjectid = new Guid(context.OutputParameters["regardingobjectid"].ToString());
                        string regardingobjectidType = "incident";

                        labor["ritsc_Case"] =
                        new EntityReference(regardingobjectidType, regardingobjectid);
                    }

                    // Create the task in Microsoft Dynamics CRM.
                    tracingService.Trace("CreateLaborForPhonePlugin: Creating the task activity.");
                    service.Create(labor);
                }

                catch (FaultException<OrganizationServiceFault> ex)
                {
                    throw new InvalidPluginExecutionException("An error occurred in CreateLaborForPhonePlugin.", ex);
                }

                catch (Exception ex)
                {
                    tracingService.Trace("CreateLaborForPhonePlugin: {0}", ex.ToString());
                    throw;
                }
            }
        }
    }
}
