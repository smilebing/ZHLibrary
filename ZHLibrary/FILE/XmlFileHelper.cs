using System.IO;
using System.Text;
using System.Xml;

namespace ZHLibrary.FILE
{
    public class XmlFileHelper
    {
        /// <summary>    
        /// 通过xsd验证xml格式是否正确，正确返回空字符串，错误返回提示    
        /// </summary>    
        /// <param name="xmlFilePath">xml文件路径</param>    
        /// <param name="xsdFilePath">xsd文件路径</param>    
        /// <param name="namespaceUrl">命名空间，无则默认为null</param>    
        /// <returns></returns>    
        public static string XmlValidationByXsd(string xmlFilePath, string xsdFilePath, string namespaceUrl = null)
        {
            StringBuilder sb = new StringBuilder();
            XmlReaderSettings settings = new XmlReaderSettings();
            settings.ValidationType = ValidationType.Schema;
            settings.Schemas.Add(namespaceUrl, xsdFilePath);
            settings.ValidationEventHandler += (x, y) =>
            {
                sb.AppendFormat("{0}", y.Message);
            };

            using (XmlReader reader = XmlReader.Create(xmlFilePath, settings))
            {
                try
                {
                    while (reader.Read())
                    {
                    }
                }
                catch (XmlException ex)
                {
                    sb.AppendFormat("{0}", ex.Message);
                }
            }
            return sb.ToString();
        }

        /// <summary>
        /// 验证xml有效性
        /// </summary>
        /// <param name="xmlStream">xml数据流</param>
        /// <param name="xsdFilePath">xsd文件路径</param>
        /// <param name="namespaceUrl">命名空间，无则默认为null</param>
        /// <returns></returns>
        public static string XmlValidationByXsd(Stream xmlStream, string xsdFilePath, string namespaceUrl = null)
        {
            StringBuilder sb = new StringBuilder();
            XmlReaderSettings settings = new XmlReaderSettings();
            settings.ValidationType = ValidationType.Schema;
            settings.Schemas.Add(namespaceUrl, xsdFilePath);
            settings.ValidationEventHandler += (x, y) =>
            {
                sb.AppendFormat("{0}", y.Message);
            };

            using (XmlReader reader = XmlReader.Create(xmlStream, settings))
            {
                try
                {
                    while (reader.Read())
                    {
                    }
                }
                catch (XmlException ex)
                {
                    sb.AppendFormat("{0}", ex.Message);
                }
            }
            return sb.ToString();
        }
    }
}
