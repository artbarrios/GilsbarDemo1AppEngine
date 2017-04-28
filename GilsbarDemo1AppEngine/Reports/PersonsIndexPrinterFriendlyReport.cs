using GilsbarDemo1.Models;
using GilsbarDemo1AppEngine.Models;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GilsbarDemo1AppEngine.Reports
{
    class PersonsIndexPrinterFriendlyReport
    {
        public static void Generate(Report report, string fileSaveDirectory, Application app)
        {
            // generates the report in the specified reportFormat with the
            // specified report.Filename saves it in fileSaveDirectory and always overwrites it
            string saveFilename = Path.Combine(fileSaveDirectory.TrimEnd('\\'), report.Filename.TrimStart('\\')) + "." + report.Extension;
            // gen up the Word objects we need
            Document document = app.Documents.Add();
            // load our styles into the document
            ReportCommon.LoadDocumentStyles(document);

            try
            {
                // build the report document
                // set the document properties
                ReportCommon.SetDocumentDefaultProperties(document, app);
                // add header
                AddDocumentHeader(document);
                // add body
                AddDocumentBody(document);
                // save the document
                document.SaveAs2(saveFilename, report.SaveFormat);
                // display ready message
                AppCommon.Log(report.Name + " ready. Open at: " + AppCommon.BuildUrl(AppCommon.GetAppEngineUrl(), report.Filename + "." + report.Extension, AppCommon.GetAppEnginePort()) + " .", EventLogEntryType.Information);
            }
            catch (Exception e)
            {
                string message = AppCommon.AppendInnerExceptionMessages("PersonsIndexViewReport.Generate: " + e.Message, e);
                message += " - Filename = " + saveFilename + "";
                throw new Exception(message);
            }
            finally
            {
                // close and dispose of the writer if it exists
                document.Close(WdSaveOptions.wdDoNotSaveChanges);
            }

        } // Generate

        private static void AddDocumentHeader(Document document)
        {
            // adds the specified part to the document
            // gen up the Word objects we need
            Paragraph paragraph;

            // get a handle to the last paragraph
            paragraph = document.Paragraphs[document.Paragraphs.Count];
            paragraph.set_Style(document.Styles["Title"]);
            paragraph.Range.Text = "Persons";

            // add trailing blank line
            document.Paragraphs.Add();
            paragraph = document.Paragraphs[document.Paragraphs.Count];
            paragraph.set_Style(document.Styles["Normal"]);
            paragraph.Range.Text = "";

        } // AddDocumentHeader()

        private static void AddDocumentBody(Document document)
        {
            // adds the specified part to the document
            // gen up the Word objects we need
            Paragraph paragraph;
            Table table;

            // get the data we need to build the report
            List<Person> personsWebData = new List<Person>();
            personsWebData = Web_Data.PersonsWebData.GetPersons();

            // add paragraph and get a handle to it
            document.Paragraphs.Add();
            paragraph = document.Paragraphs[document.Paragraphs.Count];
            paragraph.set_Style(document.Styles["Normal"]);

            // add a table and get a handle to it
            document.Tables.Add(paragraph.Range, 1, 10); // 1 X count of properties
            table = document.Tables[document.Tables.Count];
            table.set_Style(document.Styles["Plain Table 2"]);

            // set column widths
            // Example: table.Columns[1].SetWidth(app.InchesToPoints(.75f), WdRulerStyle.wdAdjustSameWidth);
            // set for no in-table page break
            table.Rows[table.Rows.Count].AllowBreakAcrossPages = 0;

            // add column headers
            // Example: table.Rows[table.Rows.Count].Cells[1].Range.Text = "Subject";
            table.Rows[table.Rows.Count].Cells[1].Range.Text = "Firstname";
            table.Rows[table.Rows.Count].Cells[2].Range.Text = "Lastname";
            table.Rows[table.Rows.Count].Cells[3].Range.Text = "Email";
            table.Rows[table.Rows.Count].Cells[4].Range.Text = "Home Phone";
            table.Rows[table.Rows.Count].Cells[5].Range.Text = "Cell Phone";
            table.Rows[table.Rows.Count].Cells[6].Range.Text = "Work Phone";
            table.Rows[table.Rows.Count].Cells[7].Range.Text = "Date of Birth";
            table.Rows[table.Rows.Count].Cells[8].Range.Text = "Manager";
            table.Rows[table.Rows.Count].Cells[9].Range.Text = "Flowchart Diagram Data";
            table.Rows[table.Rows.Count].Cells[10].Range.Text = "Flowchart Diagram Data";


            // format header row
            table.Rows[table.Rows.Count].HeadingFormat = -1;
            table.Rows[table.Rows.Count].Range.set_Style(document.Styles["TableHeaderRow"]);
            table.Rows[table.Rows.Count].Range.Bold = 1;

            // add table data rows
            foreach (Person person in personsWebData)
            {
                table.Rows.Add();
                // format data row
                table.Rows[table.Rows.Count].Range.set_Style(document.Styles["TableDataRow"]);
                table.Rows[table.Rows.Count].Range.Bold = 0;
                // Example: table.Rows[table.Rows.Count].Cells[1].Range.Text = object.Name.ToString();
                table.Rows[table.Rows.Count].Cells[1].Range.Text = person.Firstname.ToString();
                table.Rows[table.Rows.Count].Cells[2].Range.Text = person.Lastname.ToString();
                table.Rows[table.Rows.Count].Cells[3].Range.Text = person.Email.ToString();
                table.Rows[table.Rows.Count].Cells[4].Range.Text = person.HomePhone.ToString();
                table.Rows[table.Rows.Count].Cells[5].Range.Text = person.CellPhone.ToString();
                table.Rows[table.Rows.Count].Cells[6].Range.Text = person.WorkPhone.ToString();
                table.Rows[table.Rows.Count].Cells[7].Range.Text = person.DateOfBirth.ToString();
                table.Rows[table.Rows.Count].Cells[8].Range.Text = person.ManagerId.ToString();
                table.Rows[table.Rows.Count].Cells[9].Range.Text = person.JobFlowchartDiagramData.ToString();
                table.Rows[table.Rows.Count].Cells[10].Range.Text = person.JobTaskFlowchartDiagramData.ToString();

            }

            // add trailing blank line
            paragraph.Range.Text += "";

        } // AddDocumentBody()

    } // class PersonsIndexPrinterFriendlyReport
}

