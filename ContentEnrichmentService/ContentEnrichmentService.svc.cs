using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Security.Permissions;
using System.Text.RegularExpressions;
using System.Web;
using System.ServiceModel;
using System.ServiceModel.Configuration;
using System.Configuration;
using System.Web.Configuration;
using System.Web.UI;
using System.ServiceModel.Security;
using System.Security;
using System.Security.Cryptography.X509Certificates;
using System.ServiceModel.Channels;
using System.ServiceModel.Security.Tokens;
using System.IdentityModel.Tokens;
using System.IO;
using Microsoft.Office.Server.Search.ContentProcessingEnrichment;
using Microsoft.Office.Server.Search.ContentProcessingEnrichment.PropertyTypes;
using System.Globalization;
using ContentEnrichment.ServiceReference1;

namespace ContentEnrichment
{

    public class ContentEnrichmentService : IContentProcessingEnrichmentService
    {
        //private variables 
        private string applicationPrefix = "";
        private string LogID = DateTime.Now.ToFileTimeUtc().ToString();
        public StringBuilder logs = new StringBuilder();

        //private const string fileNameProperty = "Filename";
        //Store managed property names to user defined strings
        private const string input = "InputManagedProperty";
        private const string Output = "OutputManagedProperty";
        private const string updatedName = "UpdatedName";
        private const string tempDirectory = @"C:\Service\LOGS";
        string clientCertSubject = string.Empty;        // Stores the values of a clientCert.
        string defaultServCertSubject = string.Empty;   // Stores the values of a ServCert.
        string endpoint = string.Empty;                 // Stores the endpoint address value.
        string endpointDNSIdentity = string.Empty;      // Stores the values of endpointDNS.

        // Defines the error code for encountering unexpected exceptions.
        private const int UnexpectedType = 1;

        private const int UnexpectedError = 2;
        //Declare processedItemHolder to return values stored in Managed Properties.
        private readonly ProcessedItem processedItemHolder = new ProcessedItem
        {
            ItemProperties = new List<AbstractProperty>()
        };

        public ProcessedItem ProcessItem(Microsoft.Office.Server.Search.ContentProcessingEnrichment.Item item)
        {

            processedItemHolder.ErrorCode = 0;
            processedItemHolder.ItemProperties.Clear();

            if (item.ItemProperties.Any(i => i.Name == input))
            {
                CustomProcessItem(item);
            }


            foreach (var _item in item.ItemProperties)
            {
                AppendToLogs(string.Format("Property Name:{0}, Value:{1}", _item.Name, _item.ObjectValue.ToString()));
            }



            return processedItemHolder;
        }

        private void CustomProcessItem(Microsoft.Office.Server.Search.ContentProcessingEnrichment.Item item)
        {

            String Service_value = ""; //Declare global variable to store value of InputManagedProperty managed property and access it during OutputManagedProperty managed property.
            string prevResinput = ""; //Declare global variable to store name of the document which is being crawled.


            try
            {
                CustomBinding binding = new CustomBinding();
                binding.Name = "ContentEnrichmentServiceSOAP";
                //Increase timeout for binding so that service expire time increases.
                binding.OpenTimeout = new TimeSpan(2, 0, 0);
                binding.CloseTimeout = new TimeSpan(2, 0, 0);
                binding.SendTimeout = new TimeSpan(2, 0, 0);
                binding.ReceiveTimeout = new TimeSpan(2, 0, 0);

                //HTTPS Transport
                HttpsTransportBindingElement transport = new HttpsTransportBindingElement();
                transport.RequireClientCertificate = true;
                //Increase the size of message which is recieved. If the size is not increased then we get error while crawling large InputManagedProperty.
                transport.MaxReceivedMessageSize = 2147483647; //This is the max. size we can give.
                transport.MaxBufferPoolSize = 2147483647;
                transport.MaxBufferSize = 2147483647;

                //Body signing asymmetric

                AsymmetricSecurityBindingElement asec = (AsymmetricSecurityBindingElement)SecurityBindingElement.CreateMutualCertificateBindingElement(MessageSecurityVersion.WSSecurity10WSTrust13WSSecureConversation13WSSecurityPolicy12BasicSecurityProfile10, false);
                asec.SetKeyDerivation(false);
                asec.AllowInsecureTransport = true;
                asec.InitiatorTokenParameters = new X509SecurityTokenParameters { InclusionMode = SecurityTokenInclusionMode.AlwaysToRecipient, RequireDerivedKeys = false };
                asec.RecipientTokenParameters = new X509SecurityTokenParameters { InclusionMode = SecurityTokenInclusionMode.Never, RequireDerivedKeys = false };
                asec.IncludeTimestamp = false;
                asec.EnableUnsecuredResponse = true;


                //Message Encoding
                TextMessageEncodingBindingElement textMessageEncoding = new TextMessageEncodingBindingElement(MessageVersion.Soap11, Encoding.UTF8);

                //binding.Elements.Add(relSess);
                binding.Elements.Add(asec);
                binding.Elements.Add(textMessageEncoding);
                binding.Elements.Add(transport);
                //Access clientCertSubject,defaultServCertSubject,endpoint,endpointDNSIdentity from web.config file.
                clientCertSubject = System.Configuration.ConfigurationManager.AppSettings["clientCertSubject"].ToString();
                defaultServCertSubject = System.Configuration.ConfigurationManager.AppSettings["defaultServCertSubject"].ToString();
                endpoint = System.Configuration.ConfigurationManager.AppSettings["endpoint"].ToString();
                endpointDNSIdentity = System.Configuration.ConfigurationManager.AppSettings["endpointDNSIdentity"].ToString();

                //Declare object of EndpointAddress class to contain endpoint url and endpointDNSIdentity.
                EndpointAddress remoteAddress = new EndpointAddress(new Uri(endpoint),
                                                                               new DnsEndpointIdentity(endpointDNSIdentity));

                //// Retrieve the Client Certificate= Sharepoint cert.////
                X509Store store = new X509Store(StoreName.My, StoreLocation.LocalMachine);
                store.Open(OpenFlags.ReadOnly);
                X509Certificate2 cert = new X509Certificate2();
                for (int i = 0; i < store.Certificates.Count; i++)
                {
                    if (store.Certificates[i].Subject == clientCertSubject)
                    {
                        cert = store.Certificates[i];
                        break;
                    }
                }

                //// Retrieve the Server cert, ExternalSoapService- Pureapp cert ////
                X509Store store2 = new X509Store(StoreName.TrustedPeople, StoreLocation.LocalMachine);
                store2.Open(OpenFlags.ReadOnly);
                X509Certificate2 cert2 = new X509Certificate2();
                for (int i = 0; i < store2.Certificates.Count; i++)
                {
                    if (store2.Certificates[i].Subject == defaultServCertSubject)
                    {
                        cert2 = store2.Certificates[i];
                        break;
                    }
                }

                /*Create object of ContentEnrichmentServiceClient to contain binding and remote address,
                object of TranslateInputManagedPropertyRequestType to send request to the web service,
                object of TranslateInputManagedPropertyResponseType to contain response from the web service. */

                ContentEnrichmentServiceClient client = new ContentEnrichmentServiceClient(binding, remoteAddress);
                TranslateInputManagedPropertyRequestType oType = new TranslateInputManagedPropertyRequestType();
                TranslateInputManagedPropertyResponseType rType = new TranslateInputManagedPropertyResponseType();

                // Iterate over each property received and locate the two properties we
                // configured the system to send.

                foreach (var property in item.ItemProperties)
                {
                    /*If current managed property is InputManagedProperty,
                    then save the value of this property to the global variable*/
                    if (property.Name.Equals(input, StringComparison.Ordinal))
                    {
                        var InputManagedProperty = property as Property<string>;
                        Service_value = InputManagedProperty.Value;
                        processedItemHolder.ItemProperties.Add(InputManagedProperty);
                    }

                    /*If current managed property is OutputManagedProperty,
                    then save the document name to a global variable.
                    send request to web service along with the InputManagedProperty value and store the result returned from the web service to OutputManagedProperty managed property. */

                    if (property.Name.Equals(Output, StringComparison.Ordinal))
                    {
                        var re = property as Property<string>;
                        prevResinput = re.Value;
                        oType.InputManagedPropertyExpression = Service_value;
                        client.Endpoint.Contract.ProtectionLevel = System.Net.Security.ProtectionLevel.Sign;
                        client.ClientCredentials.ServiceCertificate.Authentication.CertificateValidationMode = System.ServiceModel.Security.X509CertificateValidationMode.None;
                        client.ClientCredentials.ServiceCertificate.DefaultCertificate = cert2;
                        client.ClientCredentials.ClientCertificate.Certificate = cert;

                        try
                        {
                            rType = client.translateInputManagedProperty(oType);
                        }

                        /* If there is any exception while recieving the response from web service,
                        then write the name of the document to error log. */
                        catch (Exception ex)
                        {
                            if (ex != null)
                            {
                                re.Value = "null";
                                var fullFilePath = string.Join(char.ToString(Path.DirectorySeparatorChar), tempDirectory, "Exception_Status.txt");
                                System.IO.FileStream outputFile = new FileStream(fullFilePath, FileMode.Append);
                                ExceptionMessageType expNew = new ExceptionMessageType();

                                using (StreamWriter writer = new StreamWriter(outputFile))
                                {
                                    if (expNew != null)
                                    {
                                        writer.WriteLine(System.DateTime.Now + "OutputManagedProperty value not generated for document " + prevResinput + ". Error Details: - " + ex.Message + "Detailed exception is as follows : " + ex.InnerException + "SOAP Error is as follows : " + expNew.exceptionDump);
                                    }
                                    else
                                    {
                                        writer.WriteLine(System.DateTime.Now + "OutputManagedProperty value not generated for document " + prevResinput + ". Error Details: - " + ex.Message + "Detailed exception is as follows : " + ex.InnerException);
                                    }
                                }
                            }
                            else
                            {
                                var fullFilePath2 = string.Join(char.ToString(Path.DirectorySeparatorChar), tempDirectory, "Exception_Status.txt");
                                System.IO.FileStream outputFile2 = new FileStream(fullFilePath2, FileMode.Append);

                                using (StreamWriter writer = new StreamWriter(outputFile2))
                                {
                                    writer.WriteLine(System.DateTime.Now + "There is an error in content processing element");
                                }
                            }
                        }
                        try
                        {
                            re.Value = rType.translatedInputManagedProperty;
                        }
                        catch (Exception eb)
                        {
                            var fullFilePath3 = string.Join(char.ToString(Path.DirectorySeparatorChar), tempDirectory, "Exception_Status.txt");
                            System.IO.FileStream outputFile3 = new FileStream(fullFilePath3, FileMode.Append);
                            ExceptionMessageType expNew1 = new ExceptionMessageType();
                            using (StreamWriter writer = new StreamWriter(outputFile3))
                            {
                                writer.WriteLine(System.DateTime.Now + eb.Message + "Detailed exception is as follows : " + eb.InnerException + "SOAP Exception is as follows : " + expNew1.exceptionDump);

                            }
                        }

                        processedItemHolder.ItemProperties.Add(re);
                    }
                }
            }

            /*If this service encounters any exception,
            then write that to error log. */
            catch (Exception e)
            {
                TranslateInputManagedPropertyFaultType ftype = new TranslateInputManagedPropertyFaultType();
                ExceptionMessageType mtype = new ExceptionMessageType();
                //MessageBox.Show("Bye");
                //var filename = "Exception Status_" + time + ".txt";
                var fullFilePath = string.Join(char.ToString(Path.DirectorySeparatorChar), tempDirectory, "Exception_Status.txt");
                System.IO.FileStream outputFile = new FileStream(fullFilePath, FileMode.Append);

                using (StreamWriter writer = new StreamWriter(outputFile))
                {
                    writer.WriteLine(System.DateTime.Now + e.Message + "Detailed exception is as follows : " + e.InnerException);
                }


            }
        }


        #region Private Methods

        private void CreateLog(string message)
        {
            var filename = string.Format("Service_{0}_Logs_{1}.txt", applicationPrefix, DateTime.Today.ToShortDateString().Replace("/", "_"));
            var filePath = string.Join(char.ToString(Path.DirectorySeparatorChar), new string[] { @"D:\LOGS\", filename });

            using (FileStream stream = new FileStream(filePath, FileMode.Append))
            {
                using (StreamWriter writer = new StreamWriter(stream))
                {
                    writer.WriteLine(message);
                    writer.Close();
                }
                stream.Close();
            }
        }
        private void AppendToLogs(Exception exc)
        {
            logs.AppendFormat(string.Concat(new object[] { DateTime.Now, string.Format(" > UNEXPECTED EXCEPTION: {0} | ", LogID), exc.Message, "Detailed exception is as follows : ", exc.InnerException }));
        }
        private void AppendToLogs(string message)
        {
            logs.AppendFormat("Information: {3} | {0} {1} > {2}" + System.Environment.NewLine, DateTime.Now.ToShortDateString(), DateTime.Now.ToShortTimeString(), message, LogID);
        }


        #endregion

    }
}
