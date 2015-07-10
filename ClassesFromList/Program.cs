using EnvDTE;
using Microsoft.SharePoint.Client;
using System;
using System.CodeDom;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using ClassesFromList.Util;

namespace ClassesFromList
{
    class Program
    {
        static void GenerateClass()
        {
            string outputFileName = @"C:\Users\lucas.marques\Documents\Visual Studio 2013\Projects\ClassesFromList\ClassesFromList\SampleCode.cs";

            Sample sample = new Sample();
            sample.AddFields();
            sample.AddProperties();
            sample.AddMethod();
            sample.AddConstructor();
            sample.AddEntryPoint();
            sample.GenerateCSharpCode(outputFileName);
        }

        static void Main(string[] args)
        {

            string outputFileName = @"C:\Users\lucas.marques\Documents\Visual Studio 2013\Projects\ClassesFromList\ClassesFromList\SampleCode.cs";

            //string nameUser = "lucas.marques@cspconsultoria.com.br";
            //string password = "P@ssw0rdLMM02";

            //SharepointHelper helper = new SharepointHelper();
            //ClientContext ctx = helper.GetContextWithCredentials(nameUser, password);
            //Web currentWeb = helper.GetWebCurrent(ctx);

            //helper.RetrieveListFromWeb(ctx, currentWeb, helper.NameLists[0]);

            // Console.WriteLine("Generating Class");
            //GenerateClass();

            EnvDTE80.DTE2 _dte2;
            _dte2 = (EnvDTE80.DTE2)System.Runtime.InteropServices.Marshal.GetActiveObject("VisualStudio.DTE.12.0");


            Console.WriteLine(_dte2.Solution.Projects.Item(1).FullName);

            //foreach (Project proj in _dte2.Solution.Projects)
            //{
            //    if (proj.Name == "ClassesFromList")
            //    {
            //        proj.ProjectItems.AddFromFile(outputFileName);
            //        break;
            //    }
            //}

            Console.ReadKey();

        }
    }

    class SharepointHelper
    {
        private const string _SiteUrl = "https://cspconsultoriaesistemas.sharepoint.com/sites/Dev/";
        private const string _NameList = "TestList";


        public Uri SiteUrl { get; set; }
        public List<string> NameLists { get; set; }

        #region Constructors
        public SharepointHelper() : this(_SiteUrl, _NameList) { }

        public SharepointHelper(string siteUrl, List<string> nameLists)
        {
            SiteUrl = new Uri(siteUrl);
            NameLists = nameLists;
        }

        public SharepointHelper(string siteUrl, string nameList)
        {
            SiteUrl = new Uri(siteUrl);

            NameLists = new List<string>();
            NameLists.Add(nameList);
        }
        #endregion

        #region Get SPContext from many ways
        public ClientContext GetContext()
        {
            return new ClientContext(SiteUrl);
        }

        public ClientContext GetContextWithCredentials(string u, string p)
        {
            ClientContext ctx = new ClientContext(SiteUrl);

            if (!String.IsNullOrEmpty(u) && !String.IsNullOrEmpty(p))
            {
                ctx.Credentials = GetCredentials(u, p);
            }
            return ctx;
        }

        public ClientContext GetContextWithCustomCredentials(ICredentials credentials)
        {
            try
            {
                ClientContext ctx = new ClientContext(SiteUrl);

                if (credentials != null)
                {
                    ctx.Credentials = credentials;
                }

                return ctx;
            }
            catch (Exception e)
            {
                Trace.WriteLine(String.Format("Error in: {0}", e.Message));
                return null;
            }
        }
        #endregion

        public Web GetWebCurrent(ClientContext context)
        {
            Web web = context.Web;
            context.Load(web);
            context.ExecuteQuery();
            return web;
        }

        public Site GetSiteCurrent(ClientContext context)
        {
            Site web = context.Site;
            context.Load(web);
            context.ExecuteQuery();
            return web;
        }

        public ListCollection RetrieveAllListsFromWeb(ClientContext _context, Web _web)
        {
            // Retrieve all lists from the server. 
            _context.Load(_web.Lists,
                         lists => lists.Include(list => list.Title,
                                                list => list.Id,
                                                list => list.Fields,
                                                list => list.ContentTypes));
            // Execute query. 
            _context.ExecuteQuery();

            return _web.Lists;
        }

        public List RetrieveListFromWeb(ClientContext _context, Web _web, string _nameList)
        {
            List list = _web.Lists.GetByTitle(_nameList);
            _context.Load(list, l => l.Title,
                                 l => l.Id,
                                 l => l.Fields,
                                 l => l.ContentTypes);
            _context.ExecuteQuery();

            return list;
        }

        public void Get(List list)
        {
            foreach (var f in list.Fields)
            {
                FieldType fd = f.FieldTypeKind;
            }
        }

        /// <summary>
        /// Retorna uma credencial de um usuario do sharepoint online
        /// </summary>
        /// <param name="user">UserName</param>
        /// <param name="password">Password</param>
        /// <returns></returns>
        public static SharePointOnlineCredentials GetCredentials(string user, string password)
        {
            return new SharePointOnlineCredentials(user, password.ToSecureString());
        }
    }



    /// This code example creates a graph using a CodeCompileUnit and  
    /// generates source code for the graph using the CSharpCodeProvider.
    /// </summary>
    class Sample
    {
        /// <summary>
        /// Define the compile unit to use for code generation. 
        /// </summary>
        CodeCompileUnit targetUnit;

        /// <summary>
        /// The only class in the compile unit. This class contains 2 fields,
        /// 3 properties, a constructor, an entry point, and 1 simple method. 
        /// </summary>
        CodeTypeDeclaration targetClass;

        /// <summary>
        /// The name of the file to contain the source code.
        /// </summary>
        private const string outputFileName = "SampleCode.cs";

        /// <summary>
        /// Define the class.
        /// </summary>
        public Sample()
        {
            targetUnit = new CodeCompileUnit();
            System.CodeDom.CodeNamespace samples = new System.CodeDom.CodeNamespace("EFSharepoint");
            samples.Imports.Add(new CodeNamespaceImport("System"));
            targetClass = new CodeTypeDeclaration("StartPoint");
            targetClass.IsClass = true;
            targetClass.TypeAttributes = TypeAttributes.Public | TypeAttributes.Sealed;
            samples.Types.Add(targetClass);
            targetUnit.Namespaces.Add(samples);
        }

        /// <summary>
        /// Adds two fields to the class.
        /// </summary>
        public void AddFields()
        {
            // Declare the widthValue field.
            CodeMemberField widthValueField = new CodeMemberField();
            widthValueField.Attributes = MemberAttributes.Private;
            widthValueField.Name = "widthValue";
            widthValueField.Type = new CodeTypeReference(typeof(System.Double));
            widthValueField.Comments.Add(new CodeCommentStatement(
                "The width of the object."));
            targetClass.Members.Add(widthValueField);

            // Declare the heightValue field
            CodeMemberField heightValueField = new CodeMemberField();
            heightValueField.Attributes = MemberAttributes.Private;
            heightValueField.Name = "heightValue";
            heightValueField.Type =
                new CodeTypeReference(typeof(System.Double));
            heightValueField.Comments.Add(new CodeCommentStatement(
                "The height of the object."));
            targetClass.Members.Add(heightValueField);
        }
        /// <summary>
        /// Add three properties to the class.
        /// </summary>
        public void AddProperties()
        {
            // Declare the read-only Width property.
            CodeMemberProperty widthProperty = new CodeMemberProperty();
            widthProperty.Attributes =
                MemberAttributes.Public | MemberAttributes.Final;
            widthProperty.Name = "Width";
            widthProperty.HasGet = true;

            widthProperty.Type = new CodeTypeReference(typeof(System.Double));
            widthProperty.Comments.Add(new CodeCommentStatement(
                "The Width property for the object."));
            widthProperty.GetStatements.Add(new CodeMethodReturnStatement(
                new CodeFieldReferenceExpression(
                new CodeThisReferenceExpression(), "widthValue")));
            targetClass.Members.Add(widthProperty);

            // Declare the read-only Height property.
            CodeMemberProperty heightProperty = new CodeMemberProperty();
            heightProperty.Attributes =
                MemberAttributes.Public | MemberAttributes.Final;
            heightProperty.Name = "Height";
            heightProperty.HasGet = true;
            heightProperty.Type = new CodeTypeReference(typeof(System.Double));
            heightProperty.Comments.Add(new CodeCommentStatement(
                "The Height property for the object."));
            heightProperty.GetStatements.Add(new CodeMethodReturnStatement(
                new CodeFieldReferenceExpression(
                new CodeThisReferenceExpression(), "heightValue")));
            targetClass.Members.Add(heightProperty);

            // Declare the read only Area property.
            CodeMemberProperty areaProperty = new CodeMemberProperty();
            areaProperty.Attributes =
                MemberAttributes.Public | MemberAttributes.Final;
            areaProperty.Name = "Area";
            areaProperty.HasGet = true;
            areaProperty.Type = new CodeTypeReference(typeof(System.Double));
            areaProperty.Comments.Add(new CodeCommentStatement(
                "The Area property for the object."));

            // Create an expression to calculate the area for the get accessor 
            // of the Area property.
            CodeBinaryOperatorExpression areaExpression =
                new CodeBinaryOperatorExpression(
                new CodeFieldReferenceExpression(
                new CodeThisReferenceExpression(), "widthValue"),
                CodeBinaryOperatorType.Multiply,
                new CodeFieldReferenceExpression(
                new CodeThisReferenceExpression(), "heightValue"));
            areaProperty.GetStatements.Add(
                new CodeMethodReturnStatement(areaExpression));
            targetClass.Members.Add(areaProperty);
        }

        /// <summary>
        /// Adds a method to the class. This method multiplies values stored 
        /// in both fields.
        /// </summary>
        public void AddMethod()
        {
            // Declaring a ToString method
            CodeMemberMethod toStringMethod = new CodeMemberMethod();
            toStringMethod.Attributes =
                MemberAttributes.Public | MemberAttributes.Override;
            toStringMethod.Name = "ToString";
            toStringMethod.ReturnType =
                new CodeTypeReference(typeof(System.String));

            CodeFieldReferenceExpression widthReference =
                new CodeFieldReferenceExpression(
                new CodeThisReferenceExpression(), "Width");
            CodeFieldReferenceExpression heightReference =
                new CodeFieldReferenceExpression(
                new CodeThisReferenceExpression(), "Height");
            CodeFieldReferenceExpression areaReference =
                new CodeFieldReferenceExpression(
                new CodeThisReferenceExpression(), "Area");

            // Declaring a return statement for method ToString.
            CodeMethodReturnStatement returnStatement =
                new CodeMethodReturnStatement();

            // This statement returns a string representation of the width,
            // height, and area.
            string formattedOutput = "The object:" + Environment.NewLine +
                " width = {0}," + Environment.NewLine +
                " height = {1}," + Environment.NewLine +
                " area = {2}";
            returnStatement.Expression =
                new CodeMethodInvokeExpression(
                new CodeTypeReferenceExpression("System.String"), "Format",
                new CodePrimitiveExpression(formattedOutput),
                widthReference, heightReference, areaReference);
            toStringMethod.Statements.Add(returnStatement);
            targetClass.Members.Add(toStringMethod);
        }
        /// <summary>
        /// Add a constructor to the class.
        /// </summary>
        public void AddConstructor()
        {
            // Declare the constructor
            CodeConstructor constructor = new CodeConstructor();
            constructor.Attributes =
                MemberAttributes.Public | MemberAttributes.Final;

            // Add parameters.
            constructor.Parameters.Add(new CodeParameterDeclarationExpression(
                typeof(System.Double), "width"));
            constructor.Parameters.Add(new CodeParameterDeclarationExpression(
                typeof(System.Double), "height"));

            // Add field initialization logic
            CodeFieldReferenceExpression widthReference =
                new CodeFieldReferenceExpression(
                new CodeThisReferenceExpression(), "widthValue");
            constructor.Statements.Add(new CodeAssignStatement(widthReference,
                new CodeArgumentReferenceExpression("width")));
            CodeFieldReferenceExpression heightReference =
                new CodeFieldReferenceExpression(
                new CodeThisReferenceExpression(), "heightValue");
            constructor.Statements.Add(new CodeAssignStatement(heightReference,
                new CodeArgumentReferenceExpression("height")));
            targetClass.Members.Add(constructor);
        }

        /// <summary>
        /// Add an entry point to the class.
        /// </summary>
        public void AddEntryPoint()
        {
            CodeEntryPointMethod start = new CodeEntryPointMethod();
            CodeObjectCreateExpression objectCreate =
                new CodeObjectCreateExpression(
                new CodeTypeReference("CodeDOMCreatedClass"),
                new CodePrimitiveExpression(5.3),
                new CodePrimitiveExpression(6.9));

            // Add the statement:
            // "CodeDOMCreatedClass testClass = 
            //     new CodeDOMCreatedClass(5.3, 6.9);"
            start.Statements.Add(new CodeVariableDeclarationStatement(
                new CodeTypeReference("CodeDOMCreatedClass"), "testClass",
                objectCreate));

            // Creat the expression:
            // "testClass.ToString()"
            CodeMethodInvokeExpression toStringInvoke =
                new CodeMethodInvokeExpression(
                new CodeVariableReferenceExpression("testClass"), "ToString");

            // Add a System.Console.WriteLine statement with the previous 
            // expression as a parameter.
            start.Statements.Add(new CodeMethodInvokeExpression(
                new CodeTypeReferenceExpression("System.Console"),
                "WriteLine", toStringInvoke));
            targetClass.Members.Add(start);
        }
        /// <summary>
        /// Generate CSharp source code from the compile unit.
        /// </summary>
        /// <param name="filename">Output file name</param>
        public void GenerateCSharpCode(string fileName)
        {
            CodeDomProvider provider = CodeDomProvider.CreateProvider("CSharp");
            CodeGeneratorOptions options = new CodeGeneratorOptions();
            options.BracingStyle = "C";
            using (StreamWriter sourceWriter = new StreamWriter(fileName))
            {
                provider.GenerateCodeFromCompileUnit(
                    targetUnit, sourceWriter, options);

            }
        }


    }
}
