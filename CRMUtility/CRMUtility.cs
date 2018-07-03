using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using Microsoft.Crm.Sdk.Messages;
using System.Collections.Generic;
using System.Linq;
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
 

    

        public CRMUtility (string connectionString, Action<string > action)
        {
            CrmServiceClient conn = new CrmServiceClient(connectionString);
            IOrganizationService organizationService = (IOrganizationService)conn.OrganizationWebProxyClient != null
                    ? (IOrganizationService)conn.OrganizationWebProxyClient
                    : (IOrganizationService)conn.OrganizationServiceProxy;
            if (organizationService == null)
            {
                throw new Exception("cann not connect to Dynamics CRM");
            }
            CRMOrganizationService = organizationService;
            Logger = action;
        }

        public CRMUtility(IOrganizationService service, Action<string> logger)
        {
            CRMOrganizationService = service;
            Logger = logger;
        }

        public Entity RetrieveByID(string entityName, Guid id, params string[] cols)
        {
            return CRMOrganizationService.Retrieve(entityName, id, new ColumnSet( cols));
        }
       
        public Entity RetriveEntity(string entityName, string fieldName, object fieldValue, params string[] cols)
        {
            var t = new Tuple<string, object>(fieldName, fieldValue);
            return RetriveEntity(entityName, new Tuple<string, object>[] { t }, cols);
        }

        public Entity[] RetriveAllEntity(string entityName, string fieldName, object fieldValue, params string[] cols)
        {
            var t = new Tuple<string, object>(fieldName, fieldValue);
            return RetriveAllEntity(entityName, new Tuple<string, object>[] { t }, cols);
        }

        public Entity[] RetrvieAllEntitiesIn(string entityName, string fieldName, string containValue, params string[] cols)
        {
            QueryExpression query = new QueryExpression
            {
                EntityName = entityName,
                ColumnSet = new ColumnSet(cols)
            };
            query.Criteria.AddCondition(new ConditionExpression {
                AttributeName = fieldName,
                Operator = ConditionOperator.Like,
                Values = {containValue + "%"}
            });
            
            DataCollection<Entity> arr = CRMOrganizationService.RetrieveMultiple(query).Entities;
            Log("Result count: {0}", arr.Count);
            
            return arr.ToArray();
        }


        public Entity RetriveEntity(string entityName, Tuple<string, object>[] arrTuple, params string[] cols)
        {
            Entity[] arr = RetriveAllEntity(entityName, arrTuple, cols);
            Entity entity = null;
            if (arr.Length > 0)
            {
                entity = arr[0];
            }
            return entity;
        }

        public Entity[] RetriveAllActiveEntity(string entityName, params string[] cols)
        {
            QueryExpression query = new QueryExpression
            {
                EntityName = entityName,
                ColumnSet = new ColumnSet(cols),
            };
            
            //query.Criteria.AddCondition(new ConditionExpression("statuscode",ConditionOperator.Equal,0));
            DataCollection<Entity> arr = CRMOrganizationService.RetrieveMultiple(query).Entities;
            return arr.ToArray();
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
                Log("Retrieve from {0}: {1} == {2}", entityName, tuple.Item1, tuple.Item2);
            }

            DataCollection<Entity> arr = CRMOrganizationService.RetrieveMultiple(query).Entities;
            Log("Result count: {0}", arr.Count);
            return arr.ToArray();
        }

        void setStrValue(Entity entity, string fieldName, string val)
        {
            if (!string.IsNullOrWhiteSpace(val))
            {
                entity[fieldName] = val;
            }
        }
        
        public object ObjVal(Entity entity, string key)
        {
            if (entity.Contains(key) && entity[key] != null)
            {
                return entity[key];
            }
            return null;
        }
        public int? IntVal(Entity entity, string key)
        {   
            if (entity.Contains (key) && entity [key ] != null)
            {
                return (int)entity[key];
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
    
            public string GetOptionSetValueLabel(string entityName, string fieldName, int optionSetValue)
        {

            var attReq = new RetrieveAttributeRequest();
            attReq.EntityLogicalName = entityName;
            attReq.LogicalName = fieldName;
            attReq.RetrieveAsIfPublished = true;

            var attResponse = (RetrieveAttributeResponse)CRMOrganizationService.Execute(attReq);
            var attMetadata = (EnumAttributeMetadata)attResponse.AttributeMetadata;

            return attMetadata.OptionSet.Options.Where(x => x.Value == optionSetValue).FirstOrDefault().Label.UserLocalizedLabel.Label;

        }

        public string GetEntityStringValue(Entity entity, string fld )
        {
            if (!entity.Contains(fld))
            {
                return "";
            }
            object val = entity[fld];
            if (val is string)
            {
                return val as string;
            }else if (val is EntityReference)
            {
                EntityReference reference = val as EntityReference;
                return reference.Name;
            }else if (val is DateTime)
            {
                return ((DateTime)val).ToString("MM/dd/yyyy");
            }else if (val is OptionSetValue)
            {
                OptionSetValue o = val as OptionSetValue;
                return GetOptionSetValueLabel(entity.LogicalName, fld, o.Value);
            }else if (val is decimal)
            {
                return val.ToString();
            }

            return "";
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

        public Entity CreateContactIfNotFound(string firstName, string lastName, string email,string phoneNumber, string title, Entity account, bool isPrimary)
        {
            Entity contact = null;
            if (!string.IsNullOrWhiteSpace(email))
            {
                contact =  RetriveEntity("contact", "emailaddress1",email , "emailaddress1"); 
            } else
            {
                List<Tuple<string, object>> tuples = new List<Tuple<string, object>>();
                if (!string.IsNullOrWhiteSpace (firstName))
                {
                    tuples.Add(new Tuple<string, object>("firstname", firstName));
                }
                
                tuples.Add(new Tuple<string, object>("lastname", lastName));
                tuples.Add(new Tuple<string, object>("parentcustomerid", account.Id));
                contact = RetriveEntity("contact", tuples.ToArray(), "lastname");
            }

            return CreateContact(contact, firstName, lastName, email, phoneNumber,title, account,  isPrimary);
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

        public void Create(Entity entity)
        {
            CRMOrganizationService.Create(entity);
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

       
        public void GetContactById(Guid id, out string email, out string fullName)
        {
            
            Entity contact = RetrieveByID("contact", id, "emailaddress1","fullname");
            email = contact["emailaddress1"] as string;
            fullName = contact["fullname"] as string;
        }
        public Entity CreateContact(Entity contact, string firstName, string lastName, string email, string phoneNumber, string title, Entity account, bool isPrimary)
        {
            bool create_mode = false;
            if (contact == null)
            {
                contact = new Entity();
                contact.LogicalName = "contact";
                create_mode = true;
            }
            
            if (!string.IsNullOrEmpty (firstName))
            {
                 contact["firstname"] = firstName;
            }
            if (!string.IsNullOrEmpty (lastName ))
            {
                contact["lastname"] = lastName;
            }
            contact["emailaddress1"] = email;
            contact["telephone1"] = phoneNumber;
            contact["jobtitle"] = title;

            contact["parentcustomerid"] = new EntityReference("account", account.Id);
            if (create_mode)
            {
                contact.Id = CRMOrganizationService.Create(contact);
            }
            else
            {
                CRMOrganizationService.Update(contact);
            }
            if (isPrimary)
            {
                account["primarycontactid"] = new EntityReference("contact", contact.Id);
                CRMOrganizationService.Update(account);
            }
            Logger("contact created, id= " + contact.Id);
            return contact;
        }

        public Entity CreateAccountByExtIDOrCompanyName(string companyName, string City, string Address,
            string Zip, string State, string phone, string phone1, string extId, string email,
            Money revenue,params object[] custom)
        {

            Entity account = null;
            if (!string.IsNullOrWhiteSpace(extId))
            {
                account = RetrieveAccountByExtID(extId);
            }
            else if (!string.IsNullOrWhiteSpace(Address))
            { 
                account = RetrieveAccount(companyName, Address);
            }
            else
            {
                account = RetriveEntity("account", "name", companyName);
            }
        

            return CreateAccount(account, companyName, City, Address, Zip, State, phone, phone1, extId, email, revenue, custom);
        }
    

        public Entity CreateAccount(Entity account, string companyName, string City, string Address, string Zip, string State, 
            string phone, string phone1, string extId, string email, Money revenue, params object[] custom)
        {
            bool create_mode = false;
            if (account == null)
            {
                account = new Entity();
                account.LogicalName = "account";
                create_mode = true;
            }            
            account["name"] = companyName;
            setStrValue(account, "address1_line1", Address);
            setStrValue(account, "address1_city", City);
            account["address1_postalcode"] = Zip;
            account["address1_stateorprovince"] = State;
            account["telephone1"] = phone;
            account["telephone2"] = phone1;
            account["msdyn_externalaccountid"] = extId;
            account["emailaddress1"] = email;
            account["revenue"] = revenue;
            
            string fileName = null;
            foreach (object s in custom)
            {
                if (fileName == null)
                {
                    fileName = s as string;
                } 
                else
                {
                    account[fileName] = s;
                    fileName = null;
                }
            }
            
            if (create_mode)
            {
                account.Id = CRMOrganizationService.Create(account);
            }else
            {
                CRMOrganizationService.Update(account);
            }
            return account;
        }

        public Entity RetrieveAccountByExtID(string extId)
        {
            return RetriveEntity("account", "msdyn_externalaccountid", extId );
            
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

        public void Log(string formmat, params object[] vals)
        {
            Logger(string.Format(formmat, vals));
        }
    }
}
