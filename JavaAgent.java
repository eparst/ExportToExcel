//V.Veprinskiy
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.*;

import lotus.domino.*;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
public class JavaAgent extends AgentBase {

    public void NotesMain() {
    	Session session = null;

          // (Your code goes here)
          try {
              session = getSession();
//              AgentContext agentContext = session.getAgentContext();

              Database db = session.getCurrentDatabase();
              View view = db.getView("AllVoteCurrMonth");
              View vw = db.getView("setup");
              Document docSetup = vw.getDocumentByKey("excel");
              SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy.MM.dd"); //.MM.dd G 'at' HH:mm:ss z");
              String fileName = System.getProperty("java.io.tmpdir")+ "export_" + dateFormat.format( new Date()).replace(".", "_") + ".xls";
              if (!docSetup.getItemValueString("patch").equals("")){
            	  fileName = docSetup.getItemValueString("patch") + "export_" + dateFormat.format( new Date()).replace(".", "_") + ".xls"; // "D:\\Lotus.Buffer\\export2.xls";
              }    
//              System.out.println(fileName);
              docSetup.getAttachment("export.xls").extractFile(fileName);             
              final String viewName = "Результаты голосования";              
              Vector columns = view.getColumns();
              ViewEntryCollection vec = view.getAllEntries();
              FileOutputStream fileOut = new FileOutputStream(fileName);
              HSSFWorkbook workbook = new HSSFWorkbook();
              HSSFSheet spreadsheet = workbook.createSheet(viewName);
//загаловок              
              HSSFRow row0 = spreadsheet.createRow(0);
              HSSFCell cell0 = row0.createCell(1);
//              DateFormat yrOnly = new  //SimpleDateFormat("dd", "mm", "yyyy");
              Date date = new java.util.Date();
              cell0.setCellValue("Голоса от " + date.toLocaleString());
              int rowCounter = 2; //0;
/*
 * Year - Год проведения опроса (4 символа)
Month - Месяц проведения опроса (2 символа)
Author - Ответивший (до 50 символов)
Comment - Комментарии ( до 250 символов))
Value - Оценка (от 1 до 5)              
 */
   
              ViewEntry entry = vec.getFirstEntry();
              while (entry != null) {
                  if (entry.isDocument()) {
//                  if (entry.isCategory()) {
                      HSSFRow row = spreadsheet.createRow(rowCounter++);
                      int colCounter = 0;
                      Vector columnValues = entry.getColumnValues();
                      for (int i = 0; i < columnValues.size(); i++) {
                          if (columnValues.elementAt(i) != null) {
                              ViewColumn vc = (ViewColumn) columns.elementAt(i);
                              if (!vc.isIcon() && !vc.isHidden()
                                      && !vc.isResponse()) {
                                  HSSFCell cell = row.createCell(colCounter++);
                                  cell.setCellValue(columnValues.elementAt(i)
                                          .toString());
                              }
                          }
                      }
                  }
                  entry = vec.getNextEntry(entry);
              }
              
              workbook.write(fileOut);
              fileOut.flush();
              fileOut.close();             
              Document memo = db.createDocument();
              memo.replaceItemValue("Form", "memo");
              memo.replaceItemValue("SendTo", "CN=Вепринский Виталий Львович/O=PKC");//"CN=Волохова Анна Викторовна/O=PKC"); //);//docSetup.getItemValue("SignBy"));
              RichTextItem body = memo.createRichTextItem("Body");
              body.appendText("Here is excel:");
              body.addNewLine(2);
              body.embedObject(EmbeddedObject.EMBED_ATTACHMENT, null, fileName, "report");
              memo.replaceItemValue("Subject", "Excel по оценкам за прошедший месяц");
              memo.send(false);
// only for local use              Runtime.getRuntime().exec("rundll32 SHELL32.DLL,ShellExec_RunDLL \"" + fileName + "\"");

      } catch(Exception e) {
    	  session = getSession();
    	  errh.errq(e, session);
          e.printStackTrace();
      } finally {
          try {
              if (null != session) {session.recycle();}
          } catch (NotesException ignored) {
        	  ignored.printStackTrace();
        	  errh.errq(ignored, session);
        	  }
   }
}
}
