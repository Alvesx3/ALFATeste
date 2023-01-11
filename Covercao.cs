using OfficeOpenXml;
using System.Net.Http.Headers;
using ALFA.Model;
using Newtonsoft.Json;
using System.Runtime.CompilerServices;
using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using ALFA.Model.Empresa;

namespace ALFA
{
    public class Covercao
    {
        public static async Task<string> ConversaoXLSAsync()
        {
            var DeserializedClass = new Welcome();
            try
            {
                using (var client = new HttpClient())
                {
                    client.BaseAddress = new Uri("https://crm.rdstation.com/");
                    client.DefaultRequestHeaders.Accept.Clear();
                    client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                    HttpResponseMessage response = await client.GetAsync("api/v1/organizations?token=63bdce0ba593a1000ce3c38a");

                    var jsonString = await response.Content.ReadAsStringAsync();
                    var myDeserializedClass = JsonConvert.DeserializeObject<Welcome>(jsonString);
                    DeserializedClass = myDeserializedClass;
                }

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (var excelPackage = new ExcelPackage())
                {
                    excelPackage.Workbook.Properties.Author = "Alfa sistemas";
                    excelPackage.Workbook.Properties.Title = "Lista de Empresas";

                    var sheet = excelPackage.Workbook.Worksheets.Add("Empresas");
                    sheet.Name = "Empresas";

                    var i = 1;
                    var titulos = new String[] { "Empresa", "Segmento", "URL", "Resumo","Contato", "Cargo","Telefone","Email","Email" };
                    foreach (var titulo in titulos)
                    {
                        sheet.Cells[1, i++].Value = titulo;
                    }

                    i = 1;
                    var valores = new String[] {$"{DeserializedClass.Organizations.First().Name}",
                                                $"{DeserializedClass.Organizations.First().OrganizationSegments.First().Name}", 
                                                $"{DeserializedClass.Organizations.First().Url}", 
                                                $"{DeserializedClass.Organizations.First().Resume}",
                                                $"{DeserializedClass.Organizations.First().Contacts.First().Name}",
                                                $"{DeserializedClass.Organizations.First().Contacts.First().Title}",
                                                $"{DeserializedClass.Organizations.First().Contacts.First().Phones.First().PhonePhone}",
                                                $"{DeserializedClass.Organizations.First().Contacts.First().Emails.First().EmailEmail}",
                                                $"{DeserializedClass.Organizations.First().Contacts.First().Emails.Last().EmailEmail}"};
                    foreach (var valor in valores)
                    {
                        sheet.Cells[2, i++].Value = valor;
                    }
                    string path = @"C:\temp\teste.xlsx"; //Corrigir o caminho da pasta e nome do Arquivo
                    File.WriteAllBytes(path, excelPackage.GetAsByteArray());
                }

                return "sucesso!";
            }
            catch (Exception)
            {
                throw;
            }
        }
    }
}
