using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace CRM.conexao.API
{
    public class APIManager
    {
        private readonly HttpClient client;
        private const string ApiKey = "856adfb59d45471ab288e45d3e4d9a7865f9c075cc142";
        private const string BaseUrl = "http://whatsapp.petcaesecia.com.br/api/v1/wpp";

        public APIManager()
        {
            client = new HttpClient();
        }

        public async Task<Dictionary<string, (string texto, int paramsCount, string id)>> GetTemplatesAsync()
        {
            string url = $"{BaseUrl}/listarTemplates?key={ApiKey}";
            var content = new MultipartFormDataContent
        {
            { new StringContent("a86664b9-95de-4fd2-bc68-3b1e689d0a0f"), "app_id" }
        };

            HttpResponseMessage response = await client.PostAsync(url, content);
            response.EnsureSuccessStatusCode();
            string responseBody = await response.Content.ReadAsStringAsync();

            JObject jsonResponse = JObject.Parse(responseBody);
            JArray templates = (JArray)jsonResponse["data"]["templates"];

            var templateData = new Dictionary<string, (string texto, int paramsCount, string id)>();

            foreach (var template in templates)
            {
                string nome = (string)template["nome"];
                string texto = (string)template["texto"];
                int paramsCount = (int)template["params_count"];
                string id = (string)template["id"];

                templateData[nome] = (texto, paramsCount, id);
            }

            return templateData;
        }

        public async Task<bool> OptinNumeroAsync(string numero)
        {
            var formData = new MultipartFormDataContent
        {
            { new StringContent("a86664b9-95de-4fd2-bc68-3b1e689d0a0f"), "app_id" },
            { new StringContent(numero), "numero" },
            { new StringContent("true"), "optin" }
        };

            HttpResponseMessage response = await client.PostAsync($"{BaseUrl}/alterarStatusOptinNumero?key={ApiKey}", formData);
            string responseBody = await response.Content.ReadAsStringAsync();

            JObject jsonResponse = JObject.Parse(responseBody);
            string status = (string)jsonResponse["status"];
            string msg = (string)jsonResponse["data"]["msg"];

            return status == "success" && msg.Contains("A solicitação foi feita com sucesso");
        }

        public async Task<bool> EnviarLinhaAsync(string templateId, string numero, string nomeCliente, List<string> variaveis)
        {
            var formData = new MultipartFormDataContent
        {
            { new StringContent(templateId), "template_id" },
            { new StringContent(numero), "numero" },
            { new StringContent(nomeCliente), "nome_cliente" },
            { new StringContent("[" + string.Join(",", variaveis.Select(v => $"\"{v}\"")) + "]"), "variaveis" },
            { new StringContent("Pet"), "bot" },
            { new StringContent("Inicio"), "menu_bot" }
        };

            HttpResponseMessage response = await client.PostAsync($"{BaseUrl}/enviarTemplate?key={ApiKey}", formData);
            string responseBody = await response.Content.ReadAsStringAsync();

            return response.IsSuccessStatusCode;
        }
    }
}