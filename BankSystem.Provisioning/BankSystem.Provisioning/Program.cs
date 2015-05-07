
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using OfficeDevPnP.Core.Entities;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace BankSystem.Provisioning
{
    class Program
    {
        private static Web web;

        private static ClientContext context;

        private static List clientRequests;

        private static List clientApprovedRequests;

        private static List companyRequests;

        private static List companyApprovedRequests;

        private static List approversList;

        static void Main(string[] args)
        {
            //Connecto to sp site
            context = new ClientContext(ConfigurationManager.AppSettings["TenantUrl"]);

            string userName = ConfigurationManager.AppSettings["AdminUser"];

            SecureString passWord = new SecureString();

            foreach (char c in ConfigurationManager.AppSettings["Password"].ToCharArray())
            {
                passWord.AppendChar(c);
            }
            context.Credentials = new SharePointOnlineCredentials(userName, passWord);

            web = context.Web;

            context.Load(web);
            context.Load(web, w => w.Title);

            context.ExecuteQuery();

            Console.WriteLine(web.Title);


            CreateLists();

            //CreateContentTypes();

            //CreateFields();

            //AddContetTypesToLists();

            //CreateNavigation();

            //CustomizeSite();
        }

        private static void AddContetTypesToLists()
        {
            throw new NotImplementedException();
        }

        private static void CustomizeSite()
        {
            throw new NotImplementedException();
        }

        private static void CreateNavigation()
        {
            throw new NotImplementedException();
        }

        private static void CreateFields()
        {
            throw new NotImplementedException();
        }

        private static void CreateContentTypes()
        {
            throw new NotImplementedException();
        }

        private static void CreateLists()
        {
            
            //Credit requests for client
            clientRequests = CreateSingleList("Clients Requests", "Credit Requests by People", (int)ListTemplateType.GenericList);
            context.Load(clientRequests);
            context.ExecuteQuery();


            //Approved client requests
            clientApprovedRequests = CreateSingleList("Approved Clients Requests", "Credit Requests by People that are already approved", (int)ListTemplateType.GenericList);
            context.Load(clientApprovedRequests);
            context.ExecuteQuery();


            //Credit requests for company
            companyRequests = CreateSingleList("Companies Requests", "Credit Requests by Companies", (int)ListTemplateType.GenericList);
            context.Load(companyRequests);
            context.ExecuteQuery();


            //Approved company requests
            companyApprovedRequests = CreateSingleList("Approved Companies Requests", "Credit Requests by Companies that are already approved", (int)ListTemplateType.GenericList);
            context.Load(companyApprovedRequests);
            context.ExecuteQuery();

            //Approvers list
            approversList = CreateSingleList("Approvers", "Credit Requests Approvers", (int)ListTemplateType.GenericList);
            context.Load(approversList);
            context.ExecuteQuery();
        }


        //Lists methods
        private static List CreateSingleList(string ListTitle, string ListDescription, int ListTemplate)
        {
            DeleteListIfExists(ListTitle);

            ListCreationInformation creationInfo = new ListCreationInformation();
            creationInfo.Title = ListTitle;
            creationInfo.TemplateType = ListTemplate;
            creationInfo.Url = ListTitle.Replace(" ", string.Empty);

            List list;
            list = web.Lists.Add(creationInfo);
            list.Description = ListDescription;
            list.EnableVersioning = true;
            list.Update();
            context.ExecuteQuery();

            return list;
        }

        private static void DeleteListIfExists(string ListTitle)
        {
            List list = web.Lists.GetByTitle(ListTitle);
            try
            {
                list.DeleteObject();

                context.ExecuteQuery();
            }
            catch
            { }
        }

        //Content types methods
        private static void CreateSingleContentType(string CTName,
                                                   string CTDescription,
                                                   string CTGroup,
                                                   ContentType ParrentCT,
                                                   ContentTypeCollection cTypes)
        {
            DeleteCtIfExists(CTName);

            ContentTypeCreationInformation ctInfo = new ContentTypeCreationInformation()
            {
                Name = CTName,
                Description = CTDescription,
                Group = CTGroup,
                ParentContentType = ParrentCT
            };

            cTypes.Add(ctInfo);
            context.ExecuteQuery();
            web.Update();
        }

        private static void DeleteCtIfExists(string CTName)
        {
            ContentType ct = GetContentTypeByName(CTName);
            if (ct != null)
            {
                ct.DeleteObject();
                context.ExecuteQuery();
                web.Update();
            }
        }

        private static ContentType GetContentTypeByName(string contentTypeName)
        {
            if (string.IsNullOrEmpty(contentTypeName))
                throw new ArgumentNullException("contentTypeName");

            ContentTypeCollection ctCol;

            ctCol = web.ContentTypes;

            IEnumerable<ContentType> results = web.Context.LoadQuery<ContentType>(ctCol.Where(item => item.Name == contentTypeName));
            web.Context.ExecuteQuery();
            return results.FirstOrDefault();
        }


        //Fields methods
        private static Field CreateField(FieldType fType, string displayName, string internalName, string fieldGroup, IEnumerable<KeyValuePair<string, string>> additionalAttributes, bool addToDefaultView, bool required)
        {
            Field field = null;
            if (web.FieldExistsByName(internalName))
            {
                field = web.Fields.GetByInternalNameOrTitle(internalName);
                field.DeleteObject();
                context.ExecuteQuery();
            }

            FieldCreationInformation fieldCi = new FieldCreationInformation(fType)
            {
                DisplayName = displayName,
                InternalName = internalName,
                AddToDefaultView = addToDefaultView,
                Required = required,
                Id = Guid.NewGuid(),
                Group = fieldGroup,
                AdditionalAttributes = additionalAttributes
            };
            field = web.CreateField(fieldCi);
            context.Load(field);
            return field;
        }

        private static Field CreateTaxonomyField(string displayName, string internalName, string fieldGroup, bool addToDefaultView, bool required, TermStore termStore, TermSet termSet)
        {
            Field field = null;
            if (web.FieldExistsByName(internalName))
            {
                field = web.Fields.GetByInternalNameOrTitle(internalName);
                field.DeleteObject();
                context.ExecuteQuery();
            }
            TaxonomyFieldCreationInformation fieldCi = new TaxonomyFieldCreationInformation()
            {
                DisplayName = displayName,
                InternalName = internalName,
                AddToDefaultView = addToDefaultView,
                TaxonomyItem = termSet,
                Required = required,
                Id = Guid.NewGuid(),
                Group = fieldGroup
            };

            field = web.CreateTaxonomyField(fieldCi);
            context.Load(field);

            // set the SSP ID and Term Set ID on the taxonomy field
            var taxField = web.Context.CastTo<TaxonomyField>(field);
            taxField.SspId = termStore.Id;
            taxField.TermSetId = termSet.Id;
            taxField.Update();
            web.Context.ExecuteQuery();


            return taxField;
        }

        private static Field CreateChoiceField(FieldType fType, string displayName, string internalName, string fieldGroup, IEnumerable<KeyValuePair<string, string>> additionalAttributes, bool addToDefaultView, bool required, string[] choices)
        {
            Field field = null;
            if (web.FieldExistsByName(internalName))
            {
                field = web.Fields.GetByInternalNameOrTitle(internalName);
                field.DeleteObject();
                context.ExecuteQuery();
            }
            FieldCreationInformation fieldCi = new FieldCreationInformation(FieldType.Choice)
            {
                DisplayName = displayName,
                InternalName = internalName,
                AddToDefaultView = addToDefaultView,
                Required = required,
                Id = Guid.NewGuid(),
                Group = fieldGroup,
                AdditionalAttributes = additionalAttributes,

            };
            field = web.CreateField(fieldCi);

            FieldChoice fieldChoice = context.CastTo<FieldChoice>(web.Fields.GetByTitle(displayName));
            context.Load(fieldChoice);
            context.ExecuteQuery();

            fieldChoice.Choices = choices;
            fieldChoice.Update();
            context.Load(fieldChoice);
            context.ExecuteQuery();

            return fieldChoice;
        }

        private static void AddFieldToContentType(ContentType contentType, Field field, bool required, bool hidden)
        {
            if (!contentType.IsPropertyAvailable("Id"))
            {
                web.Context.Load(contentType, ct => ct.Id);
                web.Context.ExecuteQuery();
            }

            if (!field.IsPropertyAvailable("Id"))
            {
                web.Context.Load(field, f => f.Id);
                web.Context.ExecuteQuery();
            }

            // Get the field if already exists in content type, else add field to content type
            // This will help to customize (required or hidden) any pre-existing field, also to handle existing field of Parent Content type

            web.Context.Load(contentType.FieldLinks);
            web.Context.ExecuteQuery();

            FieldLink flink = contentType.FieldLinks.FirstOrDefault(fld => fld.Id == field.Id);
            if (flink == null)
            {
                FieldLinkCreationInformation fldInfo = new FieldLinkCreationInformation();
                fldInfo.Field = field;
                contentType.FieldLinks.Add(fldInfo);
                contentType.Update(true);
                web.Context.ExecuteQuery();

                flink = contentType.FieldLinks.GetById(field.Id);
            }

            if (required || hidden)
            {
                // Update FieldLink
                flink.Required = required;
                flink.Hidden = hidden;
                contentType.Update(true);
                web.Context.ExecuteQuery();
            }
        }


        //Lists and content types methods
        private static void AddContentTypeToList(ContentType ctInfo, List list, bool isHidden)
        {
            list.ContentTypesEnabled = true;
            list.Update();
            list.Context.ExecuteQuery();

            list.ContentTypes.AddExistingContentType(ctInfo);
            list.Context.ExecuteQuery();

        }

        public static void SetDefaultContentTypeToList(List list, string contentTypeId)
        {
            ContentTypeCollection ctCol = list.ContentTypes;
            list.Context.Load(ctCol);
            list.Context.ExecuteQuery();

            var ctIds = new List<ContentTypeId>();
            foreach (ContentType ct in ctCol)
            {
                ctIds.Add(ct.Id);
            }

            var newOrder = ctIds.Except(
                // remove the folder content type
                                    ctIds.Where(id => id.StringValue.StartsWith("0x012000"))
                                 )
                                 .OrderBy(x => !x.StringValue.StartsWith(contentTypeId, StringComparison.OrdinalIgnoreCase))
                                 .ToArray();
            list.RootFolder.UniqueContentTypeOrder = newOrder;

            list.RootFolder.Update();
            list.Update();
            list.Context.ExecuteQuery();
        }
    }
}
