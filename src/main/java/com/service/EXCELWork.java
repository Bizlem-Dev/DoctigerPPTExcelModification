package com.service;

import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.ResourceBundle;


import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.format.CellDateFormatter;
import org.apache.poi.ss.formula.BaseFormulaEvaluator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.json.JSONArray;
import org.json.JSONException;
import org.json.JSONObject;

import com.sun.jersey.core.util.Base64;

public class EXCELWork {

	ResourceBundle bundleststic = ResourceBundle.getBundle("config_PPTExcel");
	static FormulaEvaluator objFormulaEvaluator=null;

	public static void main(String[] args) throws JSONException {
		// TODO Auto-generated method stub
String a="{\"<<name>>\":\"05-12-19\",\r\n" + 
		"\"<<company>>\":\"Bizlem\",\r\n" + 
		"\"<<city>>\":\"mumbai\",\r\n" + 
		"\"<<country>>\":\"india\",\r\n" + 
		"\"<<myimage>>\": \"http://gpl.bluealgo.com:8085/GPLImages/floor/a1J2s000000GmcMEAS.jpg\",\r\n" + 
		"\"<<color>>\":\"Black\",\r\n" + 
		"\"<<no1>>\":\"7\",\r\n" + 
		"\"<<no2>>\":\"5\",\r\n" + 
		"\"TemplateUrl\":\"https://bluealgo.com:8083/portal/content/services/6444/G1/DocTigerAdvanced/TemplateLibrary/Test_PPT/TemplateFile/File/Test_PPT.pptx\"\r\n" + 
		"\r\n" + 
		"\r\n" + 
		"}";



JSONObject datajson =  new JSONObject(a);
String excelpath="D:\\pallavi\\ppt\\sample2.xlsx";
String savepath ="D:\\pallavi\\ppt\\OUT_VariableReplace.xlsx";
 String  result =new EXCELWork(). parseXLSX( excelpath,  datajson,  savepath);
 

		
		
	}
	
	
	public String parseXLSX(String excelpath, JSONObject datajson, String savepath) {

		JSONObject imagesobject=new JSONObject();
		JSONArray imagesArray=null;
		
		try {
			
			if(datajson.has("imagesArray")) {
				imagesArray=datajson.getJSONArray("imagesArray");
				for(int i=0; i<imagesArray.length(); i++) {
					JSONObject subobj=imagesArray.getJSONObject(i);
					String fieldname=subobj.getString("fieldname");
					String fieldvalue=subobj.getString("fieldvalue");
					imagesobject.put(fieldname, fieldvalue)	;
				}
			}
			System.out.println("imagesobject: " + imagesobject);
			
           
            	
    			Workbook excelWookBook = null;
                FileInputStream inputStream = new FileInputStream(new File(excelpath));
    			
String fileName= excelpath.substring(excelpath.lastIndexOf(".")+1);
    			if (fileName.toLowerCase().endsWith("xlsx")) {
    				excelWookBook = new XSSFWorkbook(inputStream);
    				 	objFormulaEvaluator = new XSSFFormulaEvaluator((XSSFWorkbook) excelWookBook);
    				 	for (int i = 0; i < excelWookBook.getNumberOfSheets(); i++) {
    	    				Sheet sheet = excelWookBook.getSheetAt(i);
    	    				   getSheetDataList( sheet,  datajson);

    	    			}
    				 	//objFormulaEvaluator = new XSSFFormulaEvaluator((XSSFWorkbook) excelWookBook);


    			} else {
    				if (fileName.toLowerCase().endsWith("xls")) {
    					excelWookBook = new HSSFWorkbook(inputStream);
    					 objFormulaEvaluator = new HSSFFormulaEvaluator((HSSFWorkbook) excelWookBook);
    					 for (int i = 0; i < excelWookBook.getNumberOfSheets(); i++) {
     	    				Sheet sheet = excelWookBook.getSheetAt(i);
 	    				   getSheetDataList( sheet,  datajson);

     	    			}

    				}
    			}

    			
    			inputStream.close();
    			FileOutputStream outputStream = new FileOutputStream(savepath);
    			excelWookBook.write(outputStream);
    			excelWookBook.close();
                outputStream.close();
    			excelWookBook.close();

    	
			
			
		}catch(Exception e) {
			e.printStackTrace();
			}		
	return null;
	}

	private  void getSheetDataList(Sheet sheet, JSONObject datajson) {
		List<List<String>> ret = new ArrayList<List<String>>();

		// Get the first and last sheet row number.
		int firstRowNum = sheet.getFirstRowNum();
		int lastRowNum = sheet.getLastRowNum();
		int firstCellNum = 0;
		int lastCellNum = 0;
		
		System.out.println("firstRowNum : "+firstRowNum +"lastRowNum :"+lastRowNum);
		if (lastRowNum > 0) {
			// Loop in sheet rows.
			for (int i = firstRowNum; i < lastRowNum + 1; i++) {
				// Create a String list to save column data in a row.
				//List<String> rowDataList = new ArrayList<String>();
				// Get current row object.
				Row row = sheet.getRow(i);
				if (row == null) {
					//rowDataList.add("");
					continue;
				}
				// Get first and last cell number.
				if(i == 0){
					firstCellNum = row.getFirstCellNum();
					lastCellNum = row.getLastCellNum();
					System.out.println("firstCellNum : "+firstCellNum +"lastCellNum :"+lastCellNum);

				}
				for (int j = firstCellNum; j < lastCellNum; j++) {
					Cell cell = row.getCell(j);

					if (cell == null) {
						//rowDataList.add("");
						continue;
					}
					// Get cell type.
					int cellType = cell.getCellType();
					objFormulaEvaluator.evaluateFormulaCell(row.getCell(j));

					if (cellType == CellType.NUMERIC.getCode()) {
						double numberValue = cell.getNumericCellValue();
                        DataFormatter objDefaultFormat = new DataFormatter();
						objFormulaEvaluator.evaluate(row.getCell(j));
						objFormulaEvaluator.evaluateFormulaCell(row.getCell(j));
						// BigDecimal is used to avoid double value is counted
						// use Scientific counting method.
						// For example the original double variable value is
						// 12345678, but jdk translated the value to
						// 1.2345678E7.
						if (HSSFDateUtil.isCellDateFormatted(row.getCell(j))) {							
							 if (DateUtil.isCellDateFormatted(cell)) {
								    Date date = cell.getDateCellValue();
								    String dateFmt = "";

								    if (cell.getCellStyle().getDataFormat() == 14) { //default short date without explicit formatting
								    	if( row.getCell(j).toString().contains("-") ){
								    		dateFmt = "dd-mm-yyyy";
								    	}else{
								    		dateFmt = "dd/mm/yyyy";
								    	}
								    	 //default date format for this
								    } else { //other data formats with explicit formatting
								     dateFmt = cell.getCellStyle().getDataFormatString();
								    }
								   // String value = new CellDateFormatter(dateFmt).format(date);
								   // rowDataList.add(value);

								   }

						} else {
							//String stringCellValue = objDefaultFormat.formatCellValue(row.getCell(j),objFormulaEvaluator);
							//rowDataList.add(stringCellValue);
						}

					} else if (cellType == CellType.STRING.getCode()) {
						String cellValue = cell.getStringCellValue();
System.out.println("cellValue   "+cellValue);
System.out.println("datajson   "+datajson);

						//call a function to find pattern and set value
						String newcellvalue =getNewCellValue(cellValue, datajson);
if(isNumeric(newcellvalue)) {
	cell.setCellType(Cell.CELL_TYPE_NUMERIC);
	cell.setCellValue(Integer.parseInt(newcellvalue));

}else {
	cell.setCellValue(newcellvalue);

}
						
						//rowDataList.add(cellValue);
						
					} else if (cellType == CellType.BOOLEAN.getCode()) {
						boolean numberValue = cell.getBooleanCellValue();

						String stringCellValue = String.valueOf(numberValue);

						//rowDataList.add(stringCellValue);

					} else if (cellType == CellType.BLANK.getCode()) {
						//rowDataList.add("");
						//System.out.println(rowDataList);
					}
				}

				// Add current row data list in the return list.
				//ret.add(rowDataList);
			}
		}
		//return ret;
	}
	
	
	private static String getJSONStringFromList(List<List<String>> dataTable) throws JSONException {
		String ret = "";

		if (dataTable != null) {
			int rowCount = dataTable.size();
			if (rowCount > 1) {
				JSONArray tableJsonObject = new JSONArray();
				JSONArray f = new JSONArray();
				List<String> headerRow = dataTable.get(0);
				int columnCount = headerRow.size();
				for (int i = 1; i < rowCount; i++) {
					List<String> dataRow = dataTable.get(i);
					JSONObject rowJsonObject = new JSONObject();
					try {
						for (int j = 0; j < columnCount; j++) {
							
							String columnName = headerRow.get(j);
							String columnValue = dataRow.get(j);
							rowJsonObject.put(columnName, columnValue);
							tableJsonObject.put(rowJsonObject);

						}

					} catch (Exception e) {
						f.toString();
//						continue;
						e.printStackTrace();
					}
				}
				

				JSONObject g = new JSONObject();
				g.put("data", tableJsonObject);
				// Return string format data of JSONObject object.
				ret = g.toString();

			}
		}
		return ret;
	}
	
	public String getNewCellValue( String text, JSONObject datajson) {
		
		try {
			System.out.println("text "+text);
			System.out.println("data "+datajson);
			
			if(!text.equalsIgnoreCase("") && text!="" && text!=null && !text.equalsIgnoreCase("null")) {
			System.out.println("datajson.has(text) "+datajson.has(text));
			 if ( datajson.has(text)) {
					String value = datajson.getString(text);
					System.out.println("value" + value);
					 text = text.replaceAll("(?i)" + text, value);
					System.out.println("final text 1 " + text);

				} else {
					System.out.println("in else");

					while (text.indexOf("<<") != -1 && text.indexOf(">>") != -1) {
						System.out.println("textelementstr.indexOf(\"<<\"): " + text.indexOf("<<"));
						System.out.println("textelementstr.indexOf(\">>\"): " + text.indexOf(">>"));

						String key = text.substring(text.indexOf("<<"), text.indexOf(">>") + 2);
						System.out.println("key " + key + " textelementstr " + text);
						if (datajson.has(key)) {
							System.out.println("paramsMap.get(key)" + datajson.getString(key).toString());
							text = text.replace(key, (String) datajson.getString(key));
						} else {
							text = text.replace(key, "");
						}
					}

					while (text.indexOf("<$") != -1 && text.indexOf("$>") != -1) {
						System.out.println("textelementstr.indexOf(\"<$\"): " + text.indexOf("<$"));
						System.out.println("textelementstr.indexOf(\"$>\"): " + text.indexOf("$>"));
						String key = text.substring(text.indexOf("<$"), text.indexOf("$>") + 2);
						System.out.println("key " + key + " textelementstr " + text);
						if (datajson.has(key)) {
							text = text.replace(key, (String) datajson.getString(key));
							System.out.println("key " + key + " textelementstr " + text);
						} else {
							text = text.replace(key, "");
						}
					}

					while (text.indexOf("$<{") != -1 && text.indexOf("}>$") != -1) {
						System.out.println("textelementstr.indexOf(\"$<{\"): " + text.indexOf("$<{"));
						System.out.println("textelementstr.indexOf(\"}>$\"): " + text.indexOf("}>$"));

						String key = text.substring(text.indexOf("$<{"), text.indexOf("}>$") + 2);
						System.out.println("key " + key + " textelementstr " + text);
						if (datajson.has(key)) {
							text = text.replace(key, (String) datajson.getString(key));
							System.out.println("key " + key + " textelementstr " + text);
						} else {
							text = text.replace(key, "");
						}
					}
					while (text.indexOf("${") != -1 && text.indexOf("}$") != -1) {
						System.out.println("textelementstr.indexOf(\"${\"): " + text.indexOf("${"));
						System.out.println("textelementstr.indexOf(\"}$\"): " + text.indexOf("}$"));

						String key = text.substring(text.indexOf("${"), text.indexOf("}$") + 2);
						System.out.println("key " + key + " textelementstr " + text);
						if (datajson.has(key)) {
							text = text.replace(key, datajson.getString(key));
							System.out.println("key " + key + " textelementstr " + text);
						} else {
							text = text.replace(key, "");
						}
					}
					System.out.println("final text2 " + text);
				}
			}
		}catch(Exception e) {
			//System.out.println(e.getMessage());
			e.printStackTrace();
		}
		
		return text;
	}

	
	public static boolean isNumeric(String strNum) {
	    if (strNum == null) {
	        return false;
	    }
	    try {
	        double d = Double.parseDouble(strNum);
	    } catch (NumberFormatException nfe) {
	        return false;
	    }
	    return true;
	}
	
	
}




