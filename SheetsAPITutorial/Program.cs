using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Data = Google.Apis.Sheets.v4.Data;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using Google.Apis.Util.Store;
using System.Reflection;
using System.Linq;

namespace SheetsSample
{
    public class ServiceKey
    {
        public string type { get; set; }
        public string project_id { get; set; }
        public string private_key_id { get; set; }
        public string private_key { get; set; }
        public string client_email { get; set; }
        public string client_id { get; set; }
        public string auth_uri { get; set; }
        public string token_uri { get; set; }
        public string auth_provider_x509_cert_url { get; set; }
        public string client_x509_cert_url { get; set; }
    }

    public class MyData
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string PhoneNo { get; set; }
        public int? Age { get; set; }
        public string Address { get; set; }
        public string City { get; set; }
        public string Gender { get; set; }
        public string Country { get; set; }
        public int? boom { get; set; }
        public int? sadasd { get; set; }
    }

    public class SheetsExample
    {
        static string[] Scopes = { SheetsService.Scope.SpreadsheetsReadonly, SheetsService.Scope.DriveFile, SheetsService.Scope.Spreadsheets };
        static string applicationName = "SheetsModify";
        static SheetsService service = CreateServiceForServiceCredential(applicationName);
        static string spreadsheetId = "1sxMw83gkCDvRpzTYXQOKg2VMI-LetU7INrtXgsLZ0uo";
        static string range = "Sheet1";
        static List<Object> Headers = new List<Object> ();
        public static void Main(string[] args)
        {
            
            string x = "0";

            while (x != "6")
            {
                Console.WriteLine("Press: \n1. Reading\n2. Appending\n3. Updating\n4. Clearing\n5. Add/Update Headers\n6. Exit");
                x = Console.ReadLine();
                switch (x)
                {
                    case "1":
                        Data.ValueRange readResponse = ReadSpreadsheet(spreadsheetId, range, service);
                        if (readResponse.Values == null)
                        {
                            Console.WriteLine("Sorry! No data read from this range.");
                            break;
                        }
                        Console.WriteLine("Read " + readResponse.Values.Count + " rows.");
                        break;

                    case "2":
                        Data.AppendValuesResponse appendResponse = AppendSpreadsheet(spreadsheetId, range, service, GetDataAsObject());
                        Console.WriteLine("Appended! Updated " + appendResponse.Updates.UpdatedCells + " cells.");
                        break;

                    case "3":
                        Data.UpdateValuesResponse updateResponse = UpdateSpreadsheet(spreadsheetId, range, service, GetDataAsObject());
                        Console.WriteLine("Updated " + updateResponse.UpdatedRows + " rows and " + updateResponse.UpdatedColumns + " columns.");
                        break;
                    case "4":
                        Data.ClearValuesResponse clearResponse = Clearspreadsheet(spreadsheetId, range, service);
                        Console.WriteLine("Cleared range: " + clearResponse.ClearedRange);
                        break;
                    case "5":
                        UpdateHeadersInSpreadsheet(spreadsheetId, service, GetHeaders());
                        Console.WriteLine("Updated the headers!");
                        break;
                    default:
                        Console.WriteLine("Please enter a valid key.");
                        break;
                };
            }
        }

        public static SheetsService CreateServiceForServiceCredential(string applicationName)
        {
            SheetsService service = new SheetsService(new BaseClientService.Initializer
            {
                HttpClientInitializer = GetServiceCredential(),
                ApplicationName = applicationName
            });

            return service;
        }

        //Create sheets service for user credential (i.e asks for a user consent and pops up with google login screen)
        public static SheetsService CreateServiceForUserCredential(string applicationName)
        {
            SheetsService service = new SheetsService(new BaseClientService.Initializer
            {
                HttpClientInitializer = GetUserCredential(),
                ApplicationName = applicationName
            });

            return service;
        }

        //Create sheets service for service account (i.e no popup screen and no user consent required)
        public static UserCredential GetUserCredential()
        {
            UserCredential credential;
            using (var stream = new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal);
                credPath = Path.Combine(credPath, ".credentials/sheets.googleapis.com-dotnet-quickstart.json");
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(GoogleClientSecrets.Load(stream).Secrets, Scopes, "user", CancellationToken.None, new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }
            return credential;
        }

        public static ServiceCredential GetServiceCredential()
        {
            ServiceCredential credential;
            var _pathJson = @"My Project-2764af512c37.json";
            var json = File.ReadAllText(_pathJson);
            var cr = JsonConvert.DeserializeObject<ServiceKey>(json);
            credential = new ServiceAccountCredential(new ServiceAccountCredential.Initializer(cr.client_email)
            {
                Scopes = Scopes
            }.FromPrivateKey(cr.private_key));
            
            return credential;
        }

        public static Data.AppendValuesResponse AppendSpreadsheet(string spreadsheetId, string range, SheetsService service, MyData dataObject)
        {
            Data.ValueRange requestBody = new Data.ValueRange();

            //Request Body consists of data to be appended
            requestBody.Range = range;
            requestBody.MajorDimension = "ROWS";
            requestBody.Values = ConvertObjectToList(dataObject);

            //Update Headers w.r.t the dataObject
            UpdateHeadersInSpreadsheet(spreadsheetId, service, GetHeaders());

            //Create append request
            SpreadsheetsResource.ValuesResource.AppendRequest request = service.Spreadsheets.Values.Append(requestBody, spreadsheetId, range);
            request.ValueInputOption = (SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum)2;

            Console.WriteLine("Appending to spreadsheet...");

            //Execute the request
            Data.AppendValuesResponse response = request.Execute();
            return response;
        }

        public static Data.ValueRange ReadSpreadsheet(string spreadsheetId, string range, SheetsService service)
        {
            //Create read request
            SpreadsheetsResource.ValuesResource.GetRequest request = service.Spreadsheets.Values.Get(spreadsheetId, range);

            Console.WriteLine("Reading range " + range + " from spreadsheet...");

            //Execute the request
            Data.ValueRange response = request.Execute();
            return response;
        }

        public static Data.UpdateValuesResponse UpdateSpreadsheet(string spreadsheetId, string range, SheetsService service, MyData dataObject)
        {
            Data.ValueRange requestBody = new Data.ValueRange();
            requestBody.Values = ConvertObjectToList(dataObject);

            //Create update request
            SpreadsheetsResource.ValuesResource.UpdateRequest request = service.Spreadsheets.Values.Update(requestBody, spreadsheetId, range);
            request.ValueInputOption = (SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum)2;

            Console.WriteLine("Updating range " + range + " in spreadsheet...");

            //Execute the request
            Data.UpdateValuesResponse response = request.Execute();
            return response;
        }

        public static Data.ClearValuesResponse Clearspreadsheet(string spreadsheetId, string range, SheetsService service)
        {
            Console.WriteLine("Do you want to give your own range? 1. Yes\t Any other key. No");
            string rangeChoice = Console.ReadLine();
            if (rangeChoice == "1")
            {
                range = GetRangeFromUser();
            }

            //Create a clear request
            SpreadsheetsResource.ValuesResource.ClearRequest request = service.Spreadsheets.Values.Clear(null, spreadsheetId, range);

            Console.WriteLine("Clearing range " + range + " from spreadsheet...");

            //Execute the request
            Data.ClearValuesResponse response = request.Execute();
            Headers.Clear();
            return response;
        }

        //Get test data object
        public static MyData GetDataAsObject()
        {
            var x = new MyData();
            x.Id = 2;
            x.Name = "Harsha";
            x.Country = "India";
            x.sadasd = 83;
            return x;
        }

        //Get headers of the object (property.Name of its class) as a list
        public static List<Object> GetHeaders()
        {
            PropertyInfo[] properties = typeof(MyData).GetProperties();
            List<Object> headers = new List<Object>();
            foreach (PropertyInfo property in properties)
            {
                headers.Add(property.Name);
            }
            return headers;
        }

        public static List<Object> ReadHeaders(string spreadsheetId, SheetsService service)
        {
            Data.ValueRange response = ReadSpreadsheet(spreadsheetId, "Sheet1!1:1", service);
            if (response.Values == null)
            {
                return new List<Object> ();
            }
            return response.Values[0].ToList<Object>();
        }

        //Update headers (i.e add extra headers if any in the appending object)
        public static void UpdateHeadersInSpreadsheet(string spreadsheetId, SheetsService service, List<Object> headers)
        {
            Data.ValueRange requestBody = new Data.ValueRange();

            List<Object> headersInSpreadsheet = ReadHeaders(spreadsheetId, service);
            //Check if Headers list is equal to headers

            if (headers.All(i => headersInSpreadsheet.Contains(i)))
            {
                return; 
            }

            //Update the headers
            headersInSpreadsheet = headersInSpreadsheet.Union(headers).ToList();
            requestBody.Values = new List<IList<Object>> { headersInSpreadsheet};

            //Create update request
            SpreadsheetsResource.ValuesResource.UpdateRequest request = service.Spreadsheets.Values.Update(requestBody, spreadsheetId, "Sheet1!1:1");
            request.ValueInputOption = (SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum)2;

            Console.WriteLine("Updating headers...");

            //Execute request
            Data.UpdateValuesResponse response = request.Execute();
            return;
        }

        //Get range input from user in A1 notation
        public static string GetRangeFromUser()
        {
            Console.WriteLine("Please enter range in A1 notation.");
            string range = Console.ReadLine();
            return range;
        }

        //Convert object to list of list (i.e data type of requestBody.Values)
        public static IList<IList<Object>> ConvertObjectToList(MyData dataAsObject)
        {
            IList<Object> dataAsList = new List<Object>();
            PropertyInfo[] properties = typeof(MyData).GetProperties();
            List<Object> headersFromSpreadsheet = ReadHeaders(spreadsheetId, service);
            foreach (PropertyInfo property in properties)
            {
                if (property.GetValue(dataAsObject) != null)
                {
                    dataAsList.Add(property.GetValue(dataAsObject));
                }
                else
                {
                    dataAsList.Add("");
                }
            }
            IList<IList<Object>> testData = new List<IList<Object>> { dataAsList };
            return testData;
        }
    }
}

