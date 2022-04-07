import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
 
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;

import org.json.JSONObject;

public class ExcelConvertToJson {
	
	private ObjectMapper mapper = new ObjectMapper();
	 
    public JSONObject excelToJson(File excel) {

        List <ObjectNode> expandedTerms = new ArrayList<ObjectNode>();
        FileInputStream fis = null;
        Workbook workbook = null;
        try {
            
            fis = new FileInputStream(excel);
 
            String filename = excel.getName().toLowerCase();
            if (filename.endsWith(".xls") || filename.endsWith(".xlsx")) {
                
                if (filename.endsWith(".xls")) {
                    workbook = new HSSFWorkbook(fis);
                } else {
                    workbook = new XSSFWorkbook(fis);
                }
 
                // Reading each sheet one by one
                for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                	
                	Sheet sheet = workbook.getSheetAt(i);
                    Iterator<Row> sheetIterator = sheet.iterator();

                    while (sheetIterator.hasNext()) {

                        Row currentRow = sheetIterator.next();
          
                        ArrayNode sheetArray = mapper.createArrayNode();
                        ObjectNode expandedItems = mapper.createObjectNode();
                       	
                        Iterator<Cell> cells = currentRow.cellIterator();
                        while (cells.hasNext()) {
                        	Cell cell = cells.next();
                        		
                        	if (cell != null) {
                               if (cell.getCellTypeEnum() == CellType.STRING) {
                                  	sheetArray.add(cell.getStringCellValue());
                               } else if (cell.getCellTypeEnum() == CellType.NUMERIC) {
                                   	sheetArray.add(cell.getNumericCellValue());
                               } else if (cell.getCellTypeEnum() == CellType.BOOLEAN) {
                                   	sheetArray.add(cell.getBooleanCellValue());
                               } else if (cell.getCellTypeEnum() == CellType.BLANK) {
                                   	sheetArray.add("");
                               }
                            } else {
                                	sheetArray.add("");
                            }
                        }
                     	expandedItems.set("expanded_terms",sheetArray);
                    	expandedTerms.add(expandedItems);           	
                     }      
                }
                
                String expansion = "{ \"expansions\" : " + expandedTerms.toString() + "}";
                JSONObject json = new JSONObject(expansion);
                return json;   
            
             } else {
                throw new IllegalArgumentException("File format not supported.");
            }
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (workbook != null) {
                try {
                    workbook.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
            if (fis != null) {
                try {
                    fis.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
 
        }
        return null;
    }
 

	public static void main(String[] args) {

        File excel = new File("C:/Users/MohammadSameer/workspace/JARS/synonymsExcel.xlsx");
        ExcelConvertToJson converter = new ExcelConvertToJson();
        JSONObject data = converter.excelToJson(excel);
        System.out.println(data);
    }
}
