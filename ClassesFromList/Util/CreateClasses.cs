using System;
using System.CodeDom;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ClassesFromList.Util
{

    internal class TupleForFields<T1, T2, T3>
    {
        private readonly T1 _nameField;
        private readonly T2 _typeField;
        private readonly T3 _commentsField;

        public TupleForFields(T1 nameField, T2 typeField, T3 commentsField)
        {
            this._nameField = nameField;
            this._typeField = typeField;
            this._commentsField = commentsField;
        }

        public T1 NameField { get { return _nameField; } }

        public T2 TypeField { get { return _typeField; } }

        public T3 CommentsField { get { return _commentsField; } }
    }

    public static class Fields
    {
        public static TupleForFields<T1, T2, T3> Create<T1, T2, T3>(T1 nameField, T2 typeField, T3 commentsField)
        {
            return new TupleForFields<T1, T2, T3>(nameField, typeField, commentsField);
        }
    }

    public class CreateClass
    {
        /// <summary>
        /// The only class in the compile unit.
        /// </summary>
        private CodeTypeDeclaration _targetClass;
        public CodeTypeDeclaration TargetClass
        {
            get
            {
                return _targetClass;
            }
        }

        private string _nameClass;
        public string NameClass
        {
            get { return _nameClass; }
        }

        public CreateClass(string nameClass)
        {
            if (!String.IsNullOrEmpty(nameClass) && !String.IsNullOrWhiteSpace(nameClass))
            {
                _nameClass = nameClass;

                _targetClass = new CodeTypeDeclaration(_nameClass);
                _targetClass.IsClass = true;
                _targetClass.TypeAttributes = TypeAttributes.Public;
            }
            else
            {
                throw new Exception("Error: class need a name and it can't be white!");
            }
        }

        public void AddField(string nameField, Type typeField, string commentsField = null)
        {
            CodeMemberField valueField = new CodeMemberField();
            valueField.Attributes = MemberAttributes.Private;
            valueField.Name = nameField;
            valueField.Type = new CodeTypeReference(typeField);
            if (commentsField != null)
                valueField.Comments.Add(new CodeCommentStatement(commentsField));
            _targetClass.Members.Add(valueField);
        }

        public void AddFields(List<TupleForFields<string, Type, string>> fields)
        {
            fields.AsParallel().ForAll(item =>
            {
                CodeMemberField valueField = new CodeMemberField();
                valueField.Attributes = MemberAttributes.Private;
                valueField.Name = item.NameField;
                valueField.Type = new CodeTypeReference(item.TypeField);
                if (item.CommentsField != null)
                    valueField.Comments.Add(new CodeCommentStatement(item.CommentsField));
                _targetClass.Members.Add(valueField);
            });
        }

        public void AddProperties()
        {
            throw new NotImplementedException();
        }

        public void AddConstructors()
        {
            CodeConstructor constructor = new CodeConstructor();
            constructor.Attributes = MemberAttributes.Public;
            _targetClass.Members.Add(constructor);
        }

        public void Methods()
        {
            throw new NotImplementedException();
        }

    }

    public class GenerateClasses
    {
        /// <summary>
        /// Define the compile unit to use for code generation. 
        /// </summary>
        CodeCompileUnit _targetUnit;

        List<CreateClass> _classesToGenerate;

        private static const string Namespace = "ClassesFromListsSharepoint";


        public GenerateClasses()
        {
            _targetUnit = new CodeCompileUnit();
            _classesToGenerate = new List<CreateClass>();
        }

        public void Generate(params string[] namespaceimports)
        {
            CodeNamespace initial = new CodeNamespace(Namespace);
            if (namespaceimports != null)
            {
                for (int i = 0; i < namespaceimports.Length; i++)
                {
                    initial.Imports.Add(new CodeNamespaceImport(namespaceimports[i]));
                }
            }
        }

        private CreateClass GetNewClass(string nameClass)
        {
            CreateClass c = new CreateClass(nameClass);
            
            return c;
        }

        private void GenerateCSharpCode(string fileName)
        {
            CodeDomProvider provider = CodeDomProvider.CreateProvider("CSharp");
            CodeGeneratorOptions options = new CodeGeneratorOptions();
            options.BracingStyle = "C";
            using (StreamWriter sourceWriter = new StreamWriter(fileName))
            {
                provider.GenerateCodeFromCompileUnit(_targetUnit, sourceWriter, options);
            }
        }
    }
}
