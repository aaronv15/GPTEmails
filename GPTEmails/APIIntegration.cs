using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Http;
using System.Net.Http.Headers;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using System.Security.Policy;
using OpenAI_API;
using System.Windows.Forms;

namespace GPTEmails
{
    internal class APIIntegration
    {

        OpenAIAPI api;
        OpenAI_API.Chat.Conversation chat;

        public APIIntegration() 
        {
            api = new OpenAIAPI("API_KEY");
            chat = api.Chat.CreateConversation();
            chat.AppendSystemMessage("You are here to help write people emails. I already have a signature so do not append a sender. Do not include a subject in the email");
        }
        
        private async Task<String[]> innerRequest(string prompt, string selectedLanguage)
        {
            chat.AppendUserInput(prompt);
            string responseBody = await chat.GetResponseFromChatbotAsync();
            chat.AppendUserInput("write a subject for the email you just generated. write only the subject, do not preface or append it with anything. Please write it in " + selectedLanguage);
            string responseTitle = await chat.GetResponseFromChatbotAsync();
            string[] response = new string[] { responseBody, responseTitle };
            return response;
        }

        public string[] request(string prompt, string selectedLanguage)
        {
            Task<string[]> task = Task.Run(async () => await innerRequest(prompt, selectedLanguage));
            task.Wait();
            return task.Result;
        }
    }
}
