using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BankSystem.Provisioning
{
    class FieldNames
    {
    }
    class ListNames
    {
        public static string Clients = "Заявки за кредити от физически лица";
        public static string Companies = "Заявки за кредити от юридически лица";
    }

    class ListDescriptions
    {
        public static string Clients = "Списък със заявки за кредити от физически лица";
        public static string Companies = "Списък със заявки за кредити от юридически лица";
    }

    class ContentTypeNames
    {
        public static string Clients = "Clients Content type";
        public static string Companies = "Companies Content type";
    }

    class ContentTypeDescriptions
    {
        public static string Clients = "Content type for clients credits";
        public static string Companies = "Content type for companies credits";
    }
    
    class AdministrativeNames
    {
        public static string ColumnsGroup = "Bank System Columns";
        public static string CoontentTypesGroup = "Bank System Content Types";
    }

    class AdministrativeStrings
    {
        public static string ItemContentTypeId = "0x01";
    }
}
