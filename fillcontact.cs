using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace AutopopulateContactLookupbyEmail
{
    public class fillcontact : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            //ExecutionContext Object
            IPluginExecutionContext context = 
                (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));

            //OrganizationServiceFactory Object
            IOrganizationServiceFactory serviceFactory = 
                (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));

            //OrganizationService Object
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            Entity eng = null;

            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
                eng = (Entity)context.InputParameters["Target"];

            if (eng.LogicalName.ToLower() != "dght_engineeringemployeecompetencytable")
                return;

            if (!eng.Attributes.ContainsKey("dght_employeeemail"))
                return;

            var email = eng.GetAttributeValue<string>("dght_employeeemail");
           

            var fetch = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                              <entity name='contact'>
                                <attribute name='fullname' />
                                <attribute name='telephone1' />
                                <attribute name='contactid' />
                                <order attribute='fullname' descending='false' />
                                <filter type='and'>
                                  <condition attribute='emailaddress1' operator='eq' value='{0}' />
                                </filter>
                              </entity>
                            </fetch>";
            fetch = string.Format(fetch, email);
            EntityCollection ec = service.RetrieveMultiple(new FetchExpression(fetch));

            if (ec.Entities.Count < 1)
                return;

            Entity empcomp = ec.Entities[0];
            EntityReference contlookup = new EntityReference("contact", empcomp.Id);

            Entity ent = new Entity();
            ent.LogicalName = "dght_engineeringemployeecompetencytable";
            ent.Id = eng.Id;
            ent["dght_contacteect"] = contlookup;
            service.Update(ent);
        }
    }
}
