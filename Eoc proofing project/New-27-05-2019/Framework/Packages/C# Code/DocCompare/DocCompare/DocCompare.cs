using System;
using System.Activities;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using System.Data;
using MsWord = Microsoft.Office.Interop.Word;

namespace DocCompare
{
    public class DocCompare : CodeActivity
    {
        MsWord.Application wordApp = null;
        object readOnly = null;
        object missing = null;

        MsWord.Document doc1 = null;
        MsWord.Document doc2 = null;
        MsWord.Document doc = null;

            

        [Category("Input")]
        [RequiredArgument]
        public InArgument<object> FirstFilePath { get; set; }
        [Category("Input")]
        [RequiredArgument]
        public InArgument<object> SecondFilePath { get; set; }
        [Category("Input")]
        //[RequiredArgument]
        public InArgument<object> OutputPath { get; set; }



        //[Category("Output")]
        //public OutArgument<DataTable> Outtable { get; set; }
        protected override void Execute(CodeActivityContext context)
        {
            wordApp = new MsWord.Application();
            readOnly = true;
            missing = System.Reflection.Missing.Value;

            doc1 = wordApp.Documents.Open(FirstFilePath.Get(context), missing, readOnly,  missing, missing, missing, missing, missing, missing,
                        missing, missing, missing, missing, missing, missing, missing);
            
            doc2 = wordApp.Documents.Open(SecondFilePath.Get(context), missing, readOnly,  missing, missing, missing, missing, missing, missing,
                      missing, missing, missing, missing, missing, missing, missing);

            doc = wordApp.CompareDocuments(doc1, doc2, MsWord.WdCompareDestination.wdCompareDestinationNew,
                    MsWord.WdGranularity.wdGranularityWordLevel,
                    //System.Boolean CompareFormatting = true,
                    true,
                    //System.Boolean CompareCaseChanges = true, 
                    true,
                    //System.Boolean CompareWhitespace = true, 
                    true,
                    //System.Boolean CompareTables = true, 
                    true,
                    //System.Boolean CompareHeaders = true,
                    true,
                    //System.Boolean CompareFootnotes = true,
                    true,
                    //System.Boolean CompareTextboxes = true, 
                    true,
                    //System.Boolean CompareFields = true,  
                    true,
                    //System.Boolean CompareComments = true,  
                    true,
                    //System.Boolean CompareMoves = true,
                    true,
                    //System.String RevisedAuthor = "",                     
                    "",
                    //System.Boolean IgnoreAllComparisonWarnings = false 
                    false);

            // Close first document
            doc1.Close(missing, missing, missing);

            // Close second document
            doc2.Close(missing, missing, missing);

            // Show the compared document
            //   doc.Saved(@"D:\output\"); 
            // doc.Save();
            object filename = OutputPath.Get(context )+"\\Comparedoc.docx";
            doc.SaveAs2(ref filename);
            wordApp.Visible = true;
        }
    }
}
