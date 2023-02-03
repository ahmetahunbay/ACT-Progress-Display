/**
 * A modification of this program is used by some HAPA Academy faculty to track student progress of 
 * student ACT workbooks. Each student must have a formatted word document containing the student's 
 * name. The name must correlate with the name in the excel document used to track progress.
 * 
 */


package Excel.ExcelLoaderPractice;


import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.BodyElementType;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.net.URISyntaxException;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Hashtable;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Set;


public class Docs {
  //path to excel document
  public static final String EXCELDOC = "/pathToExcelDoc/ACTPrep-5.xlsx";
  
  /**
   * Iterates through excel sheet to create an arraylist of linkedHashMaps of progress checks by sheet
   * @return linkedhashmap, progress check as key, false as value
   */
  public static ArrayList<LinkedHashMap<String, Boolean>> getSheetCrit(){
    try {
      ArrayList<LinkedHashMap<String, Boolean>> sheetCrit= new ArrayList<LinkedHashMap<String, Boolean>>();
      FileInputStream file = 
          new FileInputStream(new File(EXCELDOC));
      XSSFWorkbook workbook = new XSSFWorkbook(file);
      
      for(Sheet sheet: workbook) {
        
        Iterator<Row> rowIterator = sheet.iterator();
        Row needed = rowIterator.next();
        Iterator<Cell> cellIterator = needed.cellIterator();        
        LinkedHashMap<String, Boolean> crit = new LinkedHashMap<String, Boolean>();
        
        //iterates through first row of each excel sheet
        while(cellIterator.hasNext()) {
          Cell runner = cellIterator.next();
         
          switch(runner.getCellType()) {
            case STRING :
              crit.put(runner.getStringCellValue(),false); 
            break;
            default:
              break;
          }      
        }
        sheetCrit.add(crit);
      }
      
      file.close();
      workbook.close();
      return sheetCrit;
    } catch(Exception e) {
      System.out.println(e.getMessage());
      return null;
    }
  }
  
  /**
   * gets row corresponding with student's name
   * @param name of student
   * @return index of row
   */
  public static int getNameInd(String name) {
    try {
      FileInputStream file = 
          new FileInputStream(new File(EXCELDOC));
      XSSFWorkbook workbook = new XSSFWorkbook(file);
      XSSFSheet sheet = workbook.getSheetAt(0);
      Iterator<Row> rowIterator = sheet.iterator();
      int ind =0;
      while(rowIterator.hasNext()) {
        Cell cell = rowIterator.next().getCell(0);
        if(cell != null && cell.getCellType() == CellType.STRING && cell.getStringCellValue().equals(name)){
          return ind;
        }
        ind++;
      }
      return -1;
    } catch( Exception e) {
      return -1;
    }
  }
  
  /**
   * gets a list of all student names
   * @return arraylist with student names
   */
  public static ArrayList<String> getNameList(){
    try {
      ArrayList<String> nameList = new ArrayList<String>();
      FileInputStream file = new FileInputStream(new File(EXCELDOC));
      XSSFWorkbook workbook = new XSSFWorkbook(file);
      Sheet sheet = workbook.getSheetAt(0);
      for(Row row: sheet) {
        Cell cell = row.getCell(0);
        if(cell != null && cell.getCellType() == CellType.STRING) {
          nameList.add(cell.getStringCellValue());
        }
      }
      
      return nameList;
    } catch (Exception e) {

      e.printStackTrace();
      return null;
    }
  }
  
  
  
  public static void main(String[] args) throws IOException, InvalidFormatException {
    try {
        //iterates through every student's document
        ArrayList<String> nameList = getNameList();
        for(int n = 0; n< nameList.size(); n++) {  
        String docUser = nameList.get(n);
        //gets doc of one student
        URL url = new URL("https://HAPA_CANVAS/" + docUser + "ACTPrep.docx");
     
        XWPFDocument xdoc = new XWPFDocument(url.openStream());
        List paragraphList = xdoc.getParagraphs();
        ArrayList<LinkedHashMap<String, Boolean>> sheetCrit= getSheetCrit();
        int l = 0;
        
        //sets all strike-through criteria to true
        for( int i = 0; i< paragraphList.size(); i++) {
          XWPFParagraph paragraph= (XWPFParagraph) paragraphList.get(i);
          XWPFRun run = paragraph.getRuns().get(0);
          if(!sheetCrit.get(l).containsKey(run.text())){
            continue;
          }
          for(int j = 0; j< sheetCrit.get(l).size();j++) {
            paragraph= (XWPFParagraph) paragraphList.get(i);
            run = paragraph.getRuns().get(0);
            if(run.isStrikeThrough() && sheetCrit.get(l).containsKey(run.text())) {
              sheetCrit.get(l).put(run.text(), true);
            }
            i++;      
          }
          l++;     
        }
        
        //compares excel headers to crossed, sets accordingly
        int nameInd = getNameInd(docUser);
        FileInputStream filein = new FileInputStream(new File(EXCELDOC));
        XSSFWorkbook workbook = new XSSFWorkbook(filein);
        
        for(int i = 0; i< workbook.getNumberOfSheets();i++) {
          XSSFSheet sheet = workbook.getSheetAt(i);
          Row row = sheet.getRow(nameInd);
          Object[] keys = sheetCrit.get(i).keySet().toArray();
          for(int j = 0; j<sheetCrit.get(i).size();j++) {
            Cell cell = row.createCell(j+1);
            cell.setCellValue(sheetCrit.get(i).get(keys[j]));
          }
        }
        
        
        filein.close();
        OutputStream os = new FileOutputStream(EXCELDOC);
        workbook.write(os);
        workbook.close();
        os.close();
      
      
        }
    
     } catch (Exception e){
       System.out.println(e.getMessage());
     }
        
        
   
  
  }
}
