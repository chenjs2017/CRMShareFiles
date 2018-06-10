using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using Microsoft.Crm.Sdk.Messages;

namespace CRMUtilityLib
{
    /*
     Microsoft.CrmSdk.CoreAssemblies
     Microsoft.CrmSdk.XrmTooling.CoreAssembly
     */
    public  class CRMUtility
    {
        public IOrganizationService CRMOrganizationService { get; set; }
        protected Action<string> Logger { get; set; }
        private void Log(string formmat, params object[] vals)
        {
            Logger(string.Format(formmat, vals));
        }

       
        /// <summary>
        /// Used in console application
        /// </summary>
        /// <param name="connectionString"></param>
        /// <param name="action"></param>
        public CRMUtility (string connectionString, Action<string > action)
        {
            CrmServiceClient conn = new CrmServiceClient(connectionString);
            IOrganizationService organizationService = (IOrganizationService)conn.OrganizationWebProxyClient != null
                    ? (IOrganizationService)conn.OrganizationWebProxyClient
                    : (IOrganizationService)conn.OrganizationServiceProxy;
            CRMOrganizationService = organizationService;
            Logger = action;
        }
        /// <summary>
        /// Used in plug-in, workflow
        /// </summary>
        /// <param name="service"></param>
        /// <param name="logger"></param>
        public CRMUtility(IOrganizationService service, Action<string> logger)
        {
            CRMOrganizationService = service;
            Logger = logger;
        }
       
        public Entity RetriveEntity(string entityName, string fieldName, object fieldValue, params string[] cols)
        {
            var t = new Tuple<string, object>(fieldName, fieldValue);
            return RetriveEntity(entityName, new Tuple<string, object>[] { t }, cols);
        }

        public Entity[] RetriveAllEntity(string entityName, string fieldName, string fieldValue, params string[] cols)
        {
            var t = new Tuple<string, object>(fieldName, fieldValue);
            return RetriveAllEntity(entityName, new Tuple<string, object>[] { t }, cols);
        }


        public Entity RetriveEntity(string entityName, Tuple<string, object>[] arrTuple, params string[] cols)
        {
            Entity[] arr = RetriveAllEntity(entityName, arrTuple, cols);
            Entity entity = null;
            if (arr.Length == 1)
            {
                entity = arr[0];
            }
            return entity;
        }

        public Entity[] RetriveAllEntity(string entityName, Tuple<string, object>[] arrTuple, params string[] cols)
        {

            QueryExpression query = new QueryExpression
            {
                EntityName = entityName,
                ColumnSet = new ColumnSet(cols),

            };
            foreach (Tuple<string, object> tuple in arrTuple)
            {
                query.Criteria.AddCondition(new ConditionExpression
                {
                    AttributeName = tuple.Item1,
                    Operator = ConditionOperator.Equal,
                    Values = { tuple.Item2 }
                });
                Logger(String.Format("Retrieve from {0}: {1} == {2}", entityName, tuple.Item1, tuple.Item2));
            }

            DataCollection<Entity> arr = CRMOrganizationService.RetrieveMultiple(query).Entities;
            Logger(String.Format("Result count: {0}", arr.Count));
            return arr.ToArray();
       }
        public int? IntVal(Entity entity, string Key)
        {
            Logger("entity = " + entity.LogicalName + " key = " + Key);
            if (entity.Contains (Key ) && entity [Key ] != null)
            {
                return (int)entity[Key];
            }
            return null;
        }
        public string StrVal(Entity entity, string key)
        {
            return entity.Contains(key) ? entity[key] as string : null;
        }
        
        public bool BoolVal(Entity entity, string Key)
        {
            return entity.Contains(Key) ? (bool) entity[Key] : false;
        }
    
        public bool ExitOpportunity(string name)
        {
            Entity[] arr = RetriveAllEntity("opportunity", "name", name, "name");
            return arr.Length > 0;
        }
        public void CreateOpportunity(string name, Guid accountID, Guid contactID)
        {
            Entity opp = new Entity();
            opp.LogicalName = "opportunity";
            opp["name"] = name;
            opp["customerid"] = new EntityReference("account", accountID);
            opp["parentcontactid"] = new EntityReference("contact", contactID);
            CRMOrganizationService.Create(opp);
        }

        public Entity CreateContactIfNotFound(string firstName, string lastName, string email,Guid acountID)
        {
            Entity contact = RetriveEntity("contact", "emailaddress1",email , "emailaddress1"); 
            if (contact == null )
            {
                contact = CreateContact(firstName, lastName ,email , acountID);
            }
            return contact;
        }

        public Entity CreateEntityIfNotExit(string logicalName, params object[] nameValuePair)
        {
            Entity entity = RetriveEntity(logicalName, nameValuePair[0] as string, nameValuePair[1], null);
            if (entity == null)
            {
               entity = CreateEntity (logicalName,nameValuePair);
            }
            return entity;
        }

        public void Update(Entity entity)
        {
            CRMOrganizationService.Update(entity);
        }

        public Entity CreateEntity(string logicalName, params object[]nameValuePair)
        { 
            Entity entity = new Entity();
            entity.LogicalName = logicalName;
            string fieldName = null;
            foreach (var val in nameValuePair)
            {
                if (fieldName == null)
                {
                    fieldName = val as string;
                }
                else
                {
                    entity[fieldName] = val;
                    fieldName = null;
                }
            }
            entity.Id = CRMOrganizationService.Create(entity);
            return entity;
        }

        public Entity CreateContact(string firstName, string lastName, string email, Guid accountID)
        {

            Logger("in create contact");
            Entity contact = new Entity();
            contact.LogicalName = "contact";
            if (!string.IsNullOrEmpty (firstName))
            {
                 contact["firstname"] = firstName;
            }
            if (!string.IsNullOrEmpty (lastName ))
            {
                contact["lastname"] = lastName;
            }
            contact["emailaddress1"] = email;
            contact["parentcustomerid"] = new EntityReference("account", accountID);
            contact.Id = CRMOrganizationService.Create(contact);
            Logger("contact created, id= " + contact.Id);
            return contact;
        }

        public Entity CreateAccount(string companyName,string City, string Address, string Zip, string State, string phone, string extent)
        {
            // be set for each entity.
            Logger("address=" + Address);
            Logger("zip=" + Zip);
            Logger("state=" + State);
            Logger("phone=" + phone);
            Logger("extent=" + extent);
            Entity account = new Entity();
            account.LogicalName = "account";

            account["name"] = companyName;
            account["address1_line1"] = Address;
            account["address1_city"] = City;
            account["address1_postalcode"] = Zip;
            account["address1_stateorprovince"] = State;
            if (!string.IsNullOrEmpty (extent))
            {
                phone += "-" + extent;
            }
            account["telephone1"] = phone;
            Logger("before crate account");
            account.Id = CRMOrganizationService.Create(account);
            return account;
        }
         
        public Entity  RetrieveAccount(string name, string address)
        {
            Tuple<string, object> nameFileter = new Tuple<string, object>("name", name);
            Tuple<string, object> addressFilter = new Tuple<string, object>("address1_line1", address);
            Tuple<string, object>[] arr = new Tuple<string, object>[] { nameFileter, addressFilter };
            return RetriveEntity("account", arr, "name");
        }

        public void ShowCRMVersion()
        {
            RetrieveVersionRequest versionRequest = new RetrieveVersionRequest();
            RetrieveVersionResponse versionResponse =
                (RetrieveVersionResponse)CRMOrganizationService.Execute(versionRequest);
            Logger(String.Format("Microsoft Dynamics CRM version {0}.", versionResponse.Version));
        }

        public void CreateBulkDeleteJob(string entityName, int executeTimePerday)
        {
            var req = new WhoAmIRequest();
            var res = (WhoAmIResponse) CRMOrganizationService.Execute(req);
            Guid guid = res.UserId;

            var condition = new ConditionExpression("statecode", ConditionOperator.NotNull);
            var filter = new FilterExpression();
            filter.AddCondition(condition);
            var query = new QueryExpression
            {
                EntityName = entityName,
                Distinct = false ,
                Criteria = filter
            };
            DateTime now = DateTime.Now;
            DateTime start = new DateTime(now.Year, now.Month, now.Day, now.Hour, 0, 0).AddHours(1);
            int steps = 24 * 60 / executeTimePerday;
            for (int i = 0; i < executeTimePerday; i ++)
            {
                var request = new BulkDeleteRequest
                {
                    JobName = "Bulk Delete " + entityName + start.ToString(" HH-mm-ss"),
                    QuerySet = new[] { query },
                    StartDateTime = start,
                    SendEmailNotification = false,
                    RecurrencePattern = "FREQ=DAILY;INTERVAL=1;",
                    ToRecipients = new Guid[] {guid },
                    CCRecipients = new Guid[] { }
                };
                CRMOrganizationService.Execute(request);
                start = start.AddMinutes(steps);
            }
        }
    }
}
