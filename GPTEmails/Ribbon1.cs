using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using System.Diagnostics;
using System.Threading.Tasks;
using Microsoft.VisualStudio.Tools.Applications.Deployment;
using Newtonsoft.Json;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Ribbon1();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace GPTEmails
{
    [ComVisible(true)]
    public class Ribbon1 : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;
        private Stack<string> previous;
        private Stack<string> future;
        private List<string> _signatureNames;
        private string selectedTemplate = "";
        private string selectedSignature = "";
        private string selectedLanguage = "English";

        bool on = false;

        APIIntegration apii = new APIIntegration();

        public Ribbon1()
        {
        }

        public void ToggleButton_OnAction(Office.IRibbonControl control, bool isPressed)
        {
            on = isPressed;

        }

        public void newEmail(Outlook.Inspector inspector)
        {
            if (inspector.CurrentItem is Outlook.MailItem mailItem && mailItem.EntryID == null && on)
            {
                mailItem.Body = string.Empty;
                mailItem.HTMLBody = string.Empty;
            }
        }

        private string queryBuilder(string emailBody, string title)
        {
            string prompt = "";
            switch (selectedTemplate)
            {
                case "Prettify":
                    prompt = "Please format this email nicely and fix any spelling errors, do not change the email otherwise. Do not add a signature or regards at the end. here is the email: " + emailBody + ". Here is the Title: " + title + ". Please write this in " + selectedLanguage;
                    break;
                case "Workplace":
                    prompt = "Please format this email so that it is appropriate for work, you may change the email slightly. Do not add a signature or regards at the end. here is the email: " + emailBody + ". Here is the Title: " + title + ". Please write this in " + selectedLanguage;
                    break;
                case "Fancy":
                    prompt = "please wrewrite this email in as fancily as you can. Do not add a signature or regards at the end. here is the email: " + emailBody + ". Here is the Title: " + title + ". Please write this in " + selectedLanguage;
                    break;
                case "Child":
                    prompt = "Please wrewrite this email as if it was written by a child. Do not add a signature or regards at the end. here is the email: " + emailBody + ". Here is the Title: " + title + ". Please write this in " + selectedLanguage;
                    break;
                case "Prompt":
                    prompt = "Do not add a signature or regards at the end unless otherwise specified. Now ";
                    break;
            }
            return prompt;
        }

        private string[] useApii(string prompt)
        {
            string[] result = apii.request(prompt, selectedLanguage);
            return result;
        }

        public void RunButton_OnAction(Office.IRibbonControl control)
        {
            if (selectedTemplate == "")
            {
                loadDefault(true, false);
            }
            string[] email = getEmail();
            string prompt = queryBuilder(email[0], email[1]);
            string[] output = useApii(prompt);
            replaceEmail(output);
            
        }

        //private void manageStack(int status)
        //{
        //    if (status == 0)
        //    {
        //        previous.Push(getEmail());
        //    }
        //    else if (status == 1)
        //    {
        //        future.Push(getEmail());
        //    }
        //}

        public void TemplateDropdown_OnAction(Office.IRibbonControl control, string selectedId, int selectedIndex)
        {
            selectedTemplate = selectedId;
        }

        public void LanguageDropdown_OnAction(Office.IRibbonControl control, string selectedId, int selectedIndex)
        {
            selectedLanguage = selectedId;
        }

        public void UndoButton_OnAction(object sender, RibbonControlEventArgs e)
        {
            // Your code to execute when the UndoButton is clicked
        }

        public void RedoButton_OnAction(object sender, RibbonControlEventArgs e)
        {
            // Your code to execute when the RedoButton is clicked
        }

        private void loadDefault(bool template, bool signature)
        {
            UserPreferences up = UserPrefrencesManager.LoadUserPreferences();
            if (template)
            {
                selectedTemplate = up.selectedTemplate;
            }
            if (signature)
            {
                selectedSignature = up.selectedSignature;
            }
        }

        public void LoadDefault_OnAction(object sender, RibbonControlEventArgs e)
        {
            loadDefault(true, true);
        }

        public void SaveButton_OnAction(object sender, RibbonControlEventArgs e)
        {
            UserPreferences up = new UserPreferences();
            up.selectedTemplate = selectedTemplate;
            up.selectedSignature = selectedSignature;
            UserPrefrencesManager.SaveUserPreferences(up);
        }

        #region Modify Emails

        private string[] getEmail()
        {
            Outlook.Inspector inspector = Globals.ThisAddIn.Application.ActiveInspector();
            Outlook.MailItem mailItem = inspector.CurrentItem as Outlook.MailItem;

            if (inspector != null && mailItem != null && !mailItem.Sent)
            {
                return new string[] { mailItem.Body, mailItem.Subject };
            }

            return null;
        }

        private string ConvertPlainTextToHtml(string plainText)
        {
            string html = plainText.Replace("\r\n", "<br>")
                .Replace("\n", "<br>")
                .Replace("*", "<strong>");

            return $"<p>{html}</p>";
        }

        public void replaceEmail(string[] customText)
        {
            Outlook.Inspector inspector = Globals.ThisAddIn.Application.ActiveInspector();
            Outlook.MailItem mailItem = inspector.CurrentItem as Outlook.MailItem;

            if (inspector != null && mailItem != null && !mailItem.Sent)
            {
                // Replace the plain text body
                mailItem.Subject = customText[1];
                mailItem.HTMLBody = ConvertPlainTextToHtml(customText[0]);

                if (selectedSignature != "")
                {
                    addSignature();
                }
            }
        }

        #endregion

        #region Signature Methods

        private void RefreshSignatureNames()
        {
            _signatureNames = GetSignatureNames();
        }

        public int SignaturesDropDown_GetItemCount(Office.IRibbonControl control)
        {
            RefreshSignatureNames();
            return _signatureNames.Count;
        }

        public string SignaturesDropDown_GetItemLabel(Office.IRibbonControl control, int index)
        {
            return _signatureNames[index];
        }

        public void SignatureDropDown_OnAction(Office.IRibbonControl control, string selectedId, int selectedIndex)
        {
            selectedSignature = _signatureNames[selectedIndex];
            // Handle the selection of a signature from the dropdown
        }

        private string AddImageAttachment(Outlook.MailItem mailItem, string imagePath)
        {
            if (File.Exists(imagePath))
            {
                Outlook.Attachment attachment = mailItem.Attachments.Add(imagePath, Outlook.OlAttachmentType.olByValue, 0, Path.GetFileName(imagePath));
                attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "cid:" + attachment.FileName);
                return "cid:" + attachment.FileName;
            }

            return string.Empty;
        }

        private List<string> GetSignatureNames()
        {
            // Get the current user's signature folder
            string appDataFolder = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string signatureFolderPath = Path.Combine(appDataFolder, "Microsoft", "Signatures");

            // Get the list of signature files
            DirectoryInfo signatureFolder = new DirectoryInfo(signatureFolderPath);
            FileInfo[] signatureFiles = signatureFolder.GetFiles("*.htm", SearchOption.TopDirectoryOnly);

            List<string> signatureNames = new List<string>();
            foreach (FileInfo signatureFile in signatureFiles)
            {
                string signatureName = Path.GetFileNameWithoutExtension(signatureFile.Name);
                signatureNames.Add(signatureName);
            }

            return signatureNames;
        }

        private string GetSignatureContent(string signatureName)
        {
            string appDataFolder = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string signatureFolderPath = Path.Combine(appDataFolder, "Microsoft", "Signatures");
            string signatureFilePath = Path.Combine(signatureFolderPath, signatureName + ".htm");

            if (File.Exists(signatureFilePath))
            {
                string signatureContent = File.ReadAllText(signatureFilePath);
                return signatureContent;
            }

            return string.Empty;
        }

        private string GetSignatureFolderPath()
        {
            string appDataFolder = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string signatureFolderPath = Path.Combine(appDataFolder, "Microsoft", "Signatures");
            return signatureFolderPath;
        }

        private void addSignature()
        {
            // Get the signature content
            string signatureContent = GetSignatureContent(selectedSignature);

            // Get the active Inspector and MailItem
            Outlook.Application application = Globals.ThisAddIn.Application;
            Outlook.Inspector activeInspector = application.ActiveInspector();

            if (activeInspector != null)
            {
                object currentItem = activeInspector.CurrentItem;
                if (currentItem is Outlook.MailItem mailItem)
                {
                    // Extract image paths from the signature content
                    string signatureFolderPath = GetSignatureFolderPath();
                    Regex imgRegex = new Regex("src=\"(?<imgPath>[^\"]*)\"", RegexOptions.IgnoreCase);
                    MatchCollection matches = imgRegex.Matches(signatureContent);

                    // Replace local image paths with Content-IDs in the signature content
                    Dictionary<string, string> imgPathToContentIdMap = new Dictionary<string, string>();
                    foreach (Match match in matches)
                    {
                        string imgPath = match.Groups["imgPath"].Value;
                        string imgFullPath = Path.Combine(signatureFolderPath, imgPath);
                        string contentId = AddImageAttachment(mailItem, imgFullPath);
                        imgPathToContentIdMap[imgPath] = contentId;
                    }

                    signatureContent = imgRegex.Replace(signatureContent, match => $"src=\"{imgPathToContentIdMap[match.Groups["imgPath"].Value]}\"");

                    // Append the signature content to the email body
                    if (mailItem.BodyFormat == Outlook.OlBodyFormat.olFormatHTML)
                    {
                        mailItem.HTMLBody += signatureContent;
                    }
                    else
                    {
                        // Convert the HTML signature to plain text
                        string plainTextSignature = Regex.Replace(signatureContent, "<[^>]*>", string.Empty);
                        mailItem.Body += plainTextSignature;
                    }
                }
            }
        }

        #endregion

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("GPTEmails.Ribbon1.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
            Globals.ThisAddIn.Ribbon1 = this;
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
