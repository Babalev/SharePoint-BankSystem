
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using OfficeDevPnP.Core;
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

        private static List companyRequests;

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

            CreateContentTypes();

            CreateFields();

            AddContetTypesToLists();

            CreateNavigation();

            //CustomizeSite();

            CreateApproversGroups();

        }

        private static void CreateApproversGroups()
        {
            GroupCreationInformation clientsApproversGroupInfo = new GroupCreationInformation()
            {
                Description = "Clients Credits Approvers group",
                Title = "Clients Credits Approvers"
            };

            GroupCreationInformation companiesApproversGroupInfo = new GroupCreationInformation()
            {
                Description = "Companies Credits Approvers group",
                Title = "Companies Credits Approvers"
            };

            web.SiteGroups.Add(clientsApproversGroupInfo);
            web.SiteGroups.Add(companiesApproversGroupInfo);
            context.ExecuteQuery();
        }

        private static void AddContetTypesToLists()
        {
            AddContentTypeToList(GetContentTypeByName(ContentTypeNames.Clients), clientRequests, false);
            SetDefaultContentTypeToList(clientRequests, GetContentTypeByName(ContentTypeNames.Clients).Id.ToString());

            AddContentTypeToList(GetContentTypeByName(ContentTypeNames.Companies), companyRequests, false);
            SetDefaultContentTypeToList(companyRequests, GetContentTypeByName(ContentTypeNames.Companies).Id.ToString());
        }

        private static void CustomizeSite()
        {
            throw new NotImplementedException();
        }

        private static void CreateNavigation()
        {
            context.Load(web.Navigation.QuickLaunch);
            context.Load(web, w => w.ServerRelativeUrl);
            context.ExecuteQuery();
            web.DeleteAllQuickLaunchNodes();
            web.Update();
            context.ExecuteQuery();


            var homeNode = web.Navigation.QuickLaunch.Add(new NavigationNodeCreationInformation()
            {
                Title = "Начало",
                Url = "",
                AsLastNode = true
            });
            homeNode.Update();

            var clientsRequestsNode = web.Navigation.QuickLaunch.Add(new NavigationNodeCreationInformation()
            {
                Title = "Clients Requests",
                Url = web.ServerRelativeUrl + string.Format("/{0}/Forms/AllItems.aspx", ListNames.Clients.Replace(" ", string.Empty)),
                AsLastNode = true
            });
            clientsRequestsNode.Update();

            var companiesRequestsNode = web.Navigation.QuickLaunch.Add(new NavigationNodeCreationInformation()
            {
                Title = "Clients Requests",
                Url = web.ServerRelativeUrl + string.Format("/{0}/Forms/AllItems.aspx", ListNames.Companies.Replace(" ", string.Empty)),
                AsLastNode = true
            });
            companiesRequestsNode.Update();
        }

        private static void CreateFields()
        {
            ContentType clientsCT = GetContentTypeByName(ContentTypeNames.Clients);
            ContentType companiesCT = GetContentTypeByName(ContentTypeNames.Companies);

            KeyValuePair<string, string> singleUserMode = new KeyValuePair<string, string>("UserSelectionMode", "0");
            KeyValuePair<string, string> peopleAndGroupsMode = new KeyValuePair<string, string>("UserSelectionMode", "1");
            KeyValuePair<string, string> dateOnlyFormat = new KeyValuePair<string, string>("Format", "DateOnly");


            //Common Fields
            Field nameField = CreateField(FieldType.Text, "Име", "Name", AdministrativeNames.ColumnsGroup, null, true, true);
            AddFieldToContentType(clientsCT, nameField, true, false);
            AddFieldToContentType(companiesCT, nameField, true, false);

            Field adressField = CreateField(FieldType.Text, "Адрес", "Address", AdministrativeNames.ColumnsGroup, null, true, true);
            AddFieldToContentType(clientsCT, adressField, true, false);
            AddFieldToContentType(companiesCT, adressField, true, false);

            Field creditSubTypeField = CreateField(FieldType.Text, "Пояснение на кредитния продукт", "CreditSubType", AdministrativeNames.ColumnsGroup, null, true, true);
            AddFieldToContentType(clientsCT, creditSubTypeField, true, false);
            AddFieldToContentType(companiesCT, creditSubTypeField, true, false);

            Field creditPurposeField = CreateField(FieldType.Note, "Цел на кредита", "CreditPurpose", AdministrativeNames.ColumnsGroup, null, true, true);
            AddFieldToContentType(clientsCT, creditPurposeField, true, false);
            AddFieldToContentType(companiesCT, creditPurposeField, true, false);

            Field creditSizeField = CreateField(FieldType.Number, "Размер на кредита", "CreditSize", AdministrativeNames.ColumnsGroup, null, true, true);
            AddFieldToContentType(clientsCT, creditSizeField, true, false);
            AddFieldToContentType(companiesCT, creditSizeField, true, false);

            string[] creditCurrencyChoices = { "BGN", "EUR", "USD" };
            Field creditCurrencyField = CreateChoiceField(FieldType.Choice, "Валута на кредита", "CreditCurrency", AdministrativeNames.ColumnsGroup, null, true, true, creditCurrencyChoices);
            AddFieldToContentType(clientsCT, creditCurrencyField, true, false);
            AddFieldToContentType(companiesCT, creditCurrencyField, true, false);

            Field creditDurationField = CreateField(FieldType.Number, "Срок на кредита", "CreditDuration", AdministrativeNames.ColumnsGroup, null, true, true);
            AddFieldToContentType(clientsCT, creditDurationField, true, false);
            AddFieldToContentType(companiesCT, creditDurationField, true, false);

            Field creditTermsField = CreateField(FieldType.Note, "Условия на кредита", "CreditTerms", AdministrativeNames.ColumnsGroup, null, true, true);
            AddFieldToContentType(clientsCT, creditTermsField, true, false);
            AddFieldToContentType(companiesCT, creditTermsField, true, false);

            Field creditSecurityField = CreateField(FieldType.Note, "Обезпечение на кредита", "CreditSecurity", AdministrativeNames.ColumnsGroup, null, true, true);
            AddFieldToContentType(clientsCT, creditSecurityField, true, false);
            AddFieldToContentType(companiesCT, creditSecurityField, true, false);

            Field userPropertiesField = CreateField(FieldType.Note, "Имуществено състояние на кредитополучателя", "UserProperties", AdministrativeNames.ColumnsGroup, null, true, true);
            AddFieldToContentType(clientsCT, userPropertiesField, true, false);
            AddFieldToContentType(companiesCT, userPropertiesField, true, false);

            Field relationshipsField = CreateField(FieldType.Note, "Взаимоотношения на клиента с банката и с други банки", "Relationships", AdministrativeNames.ColumnsGroup, null, true, true);
            AddFieldToContentType(clientsCT, relationshipsField, true, false);
            AddFieldToContentType(companiesCT, relationshipsField, true, false);



            //Client fields
            Field personalNumberField = CreateField(FieldType.Text, "ЕГН", "PersonalNumber", AdministrativeNames.ColumnsGroup, null, true, true);
            AddFieldToContentType(clientsCT, personalNumberField, true, false);

            string[] marriedChoices = { "Женен", "Неженен" };
            Field marriedField = CreateChoiceField(FieldType.Choice, "Семейно положение", "Married", AdministrativeNames.ColumnsGroup, null, true, true, marriedChoices);
            AddFieldToContentType(clientsCT, marriedField, true, false);

            string[] educationChoices = { "Начално", "Средно", "Средно Специално", "Висше" };
            Field educationField = CreateChoiceField(FieldType.Text, "Образование", "Education", AdministrativeNames.ColumnsGroup, null, true, true, educationChoices);
            AddFieldToContentType(clientsCT, educationField, true, false);

            string[] creditTypeUserChoices = { "Потребителски кредит", "Жилищен кредит", "Ипотечен кредит", "Бърз стоков кредит" };
            Field creditTypeUserField = CreateChoiceField(FieldType.Text, "Кредитен продукт за физически лица", "CreditTypeUser", AdministrativeNames.ColumnsGroup, null, true, true, creditTypeUserChoices);
            AddFieldToContentType(clientsCT, creditTypeUserField, true, false);

            Field userIncomeDocumentsField = CreateField(FieldType.Text, "Удостоверение за дохода на клиента", "UserIncomeDocuments", AdministrativeNames.ColumnsGroup, null, true, true);
            AddFieldToContentType(clientsCT, userIncomeDocumentsField, true, false);



            //Company fields
            Field eikField = CreateField(FieldType.Text, "ЕИК", "EIK", AdministrativeNames.ColumnsGroup, null, true, true);
            AddFieldToContentType(companiesCT, eikField, true, false);

            string[] creditTypeCompanyChoices = { "Кредит за оборотни средства", "Инвестиционен кредит", "Ипотечен бизнес кредит", "Кредит под условиe", "Кредитни линии" };
            Field creditTypeCompanyField = CreateChoiceField(FieldType.Choice, "Кредитен Продукт", "CreditTypeCompany", AdministrativeNames.ColumnsGroup, null, true, true, creditTypeCompanyChoices);
            AddFieldToContentType(companiesCT, creditTypeCompanyField, true, false);

            Field companyHistoryField = CreateField(FieldType.Note, "Цел на кредита", "CompanyHistory", AdministrativeNames.ColumnsGroup, null, true, true);
            AddFieldToContentType(companiesCT, companyHistoryField, true, false);

            Field connectedPeopleField = CreateField(FieldType.Note, "Свързани лица", "ConnectedPeople", AdministrativeNames.ColumnsGroup, null, true, true);
            AddFieldToContentType(companiesCT, connectedPeopleField, true, false);

            Field contractorsField = CreateField(FieldType.Note, "Контрагенти", "Contractors", AdministrativeNames.ColumnsGroup, null, true, true);
            AddFieldToContentType(companiesCT, contractorsField, true, false);

            Field competitionField = CreateField(FieldType.Note, "Конкуренция", "Competition", AdministrativeNames.ColumnsGroup, null, true, true);
            AddFieldToContentType(companiesCT, competitionField, true, false);

            Field marketTrendsField = CreateField(FieldType.Note, "Пазарни тенденции", "MarketTrends", AdministrativeNames.ColumnsGroup, null, true, true);
            AddFieldToContentType(companiesCT, marketTrendsField, true, false);

        }

        private static void CreateContentTypes()
        {
            ContentTypeCollection ContentTypesCollection = web.ContentTypes;

            ContentType itemCT = ContentTypesCollection.GetById(AdministrativeStrings.ItemContentTypeId);

            //Create clients credits CT 
            CreateSingleContentType(ContentTypeNames.Clients,
                                    ContentTypeDescriptions.Clients,
                                    AdministrativeNames.CoontentTypesGroup,
                                    itemCT, ContentTypesCollection);

            //Create companies credits CT 
            CreateSingleContentType(ContentTypeNames.Companies,
                                    ContentTypeDescriptions.Companies,
                                    AdministrativeNames.CoontentTypesGroup,
                                    itemCT, ContentTypesCollection);

        }

        private static void CreateLists()
        {

            //Credit requests for client
            clientRequests = CreateSingleList(ListNames.Clients, ListDescriptions.Clients, (int)ListTemplateType.GenericList);
            context.Load(clientRequests);
            context.ExecuteQuery();

            //Credit requests for company
            companyRequests = CreateSingleList(ListNames.Companies, ListDescriptions.Companies, (int)ListTemplateType.GenericList);
            context.Load(companyRequests);
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
