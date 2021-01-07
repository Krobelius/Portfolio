using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Serialization;
using System.Text;
using System.IO;
using System.Threading.Tasks;

namespace WEReplace1._0
{
    public class PropsFields
    {
        public string XMLFileName = "sett.xml";
        public string path_value = @"null";
    }
    public class Props
    {
        public PropsFields Fields;

        public Props()
        {
            Fields = new PropsFields();
        }
        public void WriteXml()
        {
            XmlSerializer ser = new XmlSerializer(typeof(PropsFields));
            TextWriter writer = new StreamWriter(Fields.XMLFileName);
            ser.Serialize(writer, Fields);
            writer.Close();
        }
        public void ReadXml()
        {
            if (File.Exists(Fields.XMLFileName))
            {
                XmlRootAttribute xRoot = new XmlRootAttribute
                {
                    ElementName = "PropsFields",
                    IsNullable = true
                };
                XmlSerializer ser = new XmlSerializer(typeof(PropsFields),xRoot);
                TextReader reader = new StreamReader(Fields.XMLFileName);
                Fields = ser.Deserialize(reader) as PropsFields;
                reader.Close();
            }
        }
    }
}
