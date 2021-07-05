package com.example.demo;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.io.Writer;
import java.text.MessageFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Properties;
import java.util.stream.Collectors;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.xerces.xs.LSInputList;
import org.apache.xml.resolver.apps.xparse;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.example.utils.XWPFUtils;
import com.monitorjbl.xlsx.StreamingReader;

import fr.opensagres.xdocreport.core.utils.StringUtils;
import freemarker.template.Configuration;
import freemarker.template.Template;
import freemarker.template.TemplateException;
import java_cup.internal_error;

public class ExcelToWord {

	private final static Logger logger = LoggerFactory.getLogger(ExcelToWord.class);
	List<String> queryColArray;// 要抓取的欄位
	List<String> tableOutputKey;// 範本需要抓取的欄位
	List<String> table3OutputKey;// 範本output需要抓取的欄位
	List<String> table4OutputKey;
	List<String> table5OutputKey;
	List<String> table6OutputKey;
	String excelFolderPath; // excel資料夾位置
	String tempFileFolderPath; // word範本位置
	String destFileFolderPath; // word輸出資料夾位置
	String JCLNameLast;// 存放JCLName
	String systemName;
	String sheetName;
	XWPFUtils XWPFUtils = new XWPFUtils();

	// (CellIndex,HeaderName)
	Map<Integer, String> HeaderName = new HashMap<Integer, String>();

	ExcelToWord() {
		Properties pro = new Properties();
		// 設定檔位置
		String config = "config.properties";
		try {
			System.out.println("ExcelToWord 執行");
			// 讀取設定檔
			logger.info(MessageFormat.format("設定檔位置: {0}", config));
			pro.load(new FileInputStream(config));
			// 讀取資料夾位置
			excelFolderPath = pro.getProperty("excelDir");
			tempFileFolderPath = pro.getProperty("tempFile");
			destFileFolderPath = pro.getProperty("destFile");
			logger.info(" Excel資料夾位置:{}\n Word範例檔位置:{}\n 輸出資料夾位置:{}", excelFolderPath, tempFileFolderPath,
					destFileFolderPath);
			// 讀取需要抓取的欄位名稱
			queryColArray = Arrays.asList(pro.getProperty("queryColArray").split(","));
			tableOutputKey = Arrays.asList(pro.getProperty("tableOutputKey").split(","));
			table3OutputKey = Arrays.asList(pro.getProperty("table3OutputKey").split(","));
			table4OutputKey = Arrays.asList(pro.getProperty("table4OutputKey").split(","));
			table5OutputKey = Arrays.asList(pro.getProperty("table5OutputKey").split(","));
			table6OutputKey = Arrays.asList(pro.getProperty("table6OutputKey").split(","));
			logger.info("需要抓取的欄位 " + queryColArray);
		} catch (FileNotFoundException e) {
			logger.info(e.toString());
			e.printStackTrace();
		} catch (IOException e) {
			logger.info(e.toString());
			e.printStackTrace();
		} catch (Exception e) {
			logger.info(e.toString());
			e.printStackTrace();
		}
	}

	public void excelToWordStart() throws Exception {
		// 取資料夾
		File excelFolder = new File(excelFolderPath);
		logger.info("excelDir:{} 有 {} 個Excel檔案", excelFolderPath, excelFolder.list().length);

		for (File file : excelFolder.listFiles()) {
			logger.info("開始讀取 " + file.getName());
			// 讀取excel檔案
			Workbook wb = getExcelFile(file.getPath());
			if (wb == null) {
				logger.info(file.getPath() + "讀取失敗");
				throw new Exception("讀取失敗");
			}
			logger.info(file.getName() + " 讀取完成");
			// 解析Excel to List
			logger.info("開始解析 " + file.getName());
			List<Map<String, String>> excelInfoList = parseExcel(wb);
			logger.info("解析 " + file.getName() + " 完成");
			// excel檔名
			// EX: excel檔名 帳務作業流程清單(BANK)_1090430 取 帳務作業流程清單(BANK)
			systemName = file.getName();
			logger.info("預計輸出檔名 " + sheetName);
			outPutToWork(excelInfoList);
			logger.info(systemName + " 輸出完成");
		}
	}

	/**
	 * 讀取excel檔案
	 * 
	 * Workbook(Excel本體)、Sheet(內部頁面)、Row(頁面之行(橫的))、Cell(行內的元素)
	 * 
	 * 
	 * @param path excel檔案路徑
	 * @return excel內容
	 */
	public Workbook getExcelFile(String path) {
		Workbook wb = null;
		try {
			if (path == null) {
				return null;
			}
			String extString = path.substring(path.lastIndexOf(".")).toLowerCase();
			FileInputStream in = new FileInputStream(path);
			wb = StreamingReader.builder().rowCacheSize(100)// 存到記憶體行數，預設10行。
					.bufferSize(2048)// 讀取到記憶體的上限，預設1024
					.open(in);
		} catch (FileNotFoundException e) {
			logger.info(e.toString());
			e.printStackTrace();
		}

		return wb;
	}

	/**
	 * 解析Sheet
	 * 
	 * @param workbook Excel檔案
	 * @return 整個Sheet的資料
	 */
	public List<Map<String, String>> parseExcel(Workbook workbook) {
		// Sheet的資料
		List<Map<String, String>> excelDataList = new ArrayList<>();
		// 存放DNS欄位的欄位號
		int dnsIndex = 0;
		// 遍歷每一個sheet
		for (int sheetNum = 0; sheetNum < workbook.getNumberOfSheets(); sheetNum++) {
			Sheet sheet = workbook.getSheetAt(sheetNum);
			boolean rowNum = true;

			sheetName = sheet.getSheetName();

			// 開始讀取sheet
			for (Row row : sheet) {
				// 先取header
				if (rowNum) {
					for (Cell cell : row) {
//						queryColArray.get(queryColArray.indexOf("JCL_Description"));
						if (queryColArray.contains(cell.getStringCellValue())) {
//							if (cell.getStringCellValue() == "DSN") {
//								dnsIndex = cell.getColumnIndex();
//							} else {
						
							HeaderName.put(cell.getColumnIndex(), cell.getStringCellValue());
//							}
//							System.out.println(HeaderName);
						}
					}
					rowNum = false;
					continue;
				}

				/*
				 * OLD code Row firstRow = sheet.getRow(firstRowNum); if (null == firstRow) {
				 * System.out.println("解析Excel失敗"); } int rowStart = sheetNum;// 起始去掉首欄 int
				 * rowEnd = sheet.getPhysicalNumberOfRows();OLD int dnsIndex = 0; old for (Cell
				 * cell : firstRow) { if (cell.getStringCellValue().equals("DSN")) { dnsIndex =
				 * cell.getColumnIndex(); } } for (int rowNum = rowStart; rowNum < rowEnd;
				 * rowNum++) { Row row = sheet.getRow(rowNum); if (null == row) { continue; } //
				 * 解析Row的資料 excelDataList.add(convertRowToData(row, firstRow, dnsIndex)); }-
				 */
				// 解析Row的資料
				
				excelDataList.add(convertRowToData(row, dnsIndex));
//				System.out.println(excelDataList);
			}
		}
		return excelDataList;
	}

	/**
	 * 將資料重組並輸出Word
	 * 
	 * @param excelDataList 整理過的Excel檔案
	 */
	public void outPutToWork(List<Map<String, String>> excelDataList) {
		
		//篩選Delete_Reason_EXEC_DD
		Iterator<Map<String, String>> ppIterator=excelDataList.iterator();
		while(ppIterator.hasNext()) {
			Map<String, String> reMap=ppIterator.next();
			if(!reMap.get("Delete_Reason_EXEC_DD").equals("")) {
				ppIterator.remove();
			}
		}
		
		
		
		// 抓出不重複的JCL
		HashSet<String> jclKeys = new HashSet<>();
		excelDataList.forEach(cn -> {
			jclKeys.add(cn.get("JCL"));
		});
	
		int fileCount = 0;
		logger.info("預計產出 {} 個檔案", jclKeys.size());
		// 將不重複的相同JCL_NAME的資料Group to List並輸出word
		for (String classKey : jclKeys) {
			logger.info("開始輸出:" + classKey);
			List<Map<String, String>> toWordList;
			toWordList = excelDataList.stream().filter(excelModel -> excelModel.get("JCL") == classKey
						)
					.collect(Collectors.toList()); // 篩選classkey之後回傳
			// 輸出Word
//			System.out.println(toWordList);
//			System.out.println(excelDataList);
			
//			if(excelDataList.contains("Delete_Reason_EXEC_DD").equals("")) {
//				
//			}
			createWord(toWordList);
			logger.info("已輸出 JCL Name: " + classKey);
			// 輸出完之後，刪除，節省資源。
			toWordList.forEach(Item -> excelDataList.remove(Item));
			toWordList.clear();
			fileCount++;
		}
		logger.info("實際產出 {} 個檔案", fileCount);
	}

	/**
	 * 解析ROW
	 * 
	 * @param row      資料行
	 * @param firstRow 標頭
	 * @param dnsIndex Dns的列數
	 * @return 整row的欄位
	 */
	public Map<String, String> convertRowToData(Row row, int dnsIndex) {
		Map<String, String> excelDateMap = new HashMap<String, String>();

		for (Object key : HeaderName.keySet()) {
			for (Cell cell : row) {

				// 1.
				int headerNameIndex = (int) key;

				cell = row.getCell(headerNameIndex, MissingCellPolicy.CREATE_NULL_AS_BLANK);
				if (cell == null) {
					cell.setCellValue("Empty");
//			}else if(!(cell.getStringCellValue().trim().length() > 0)) {

				}
				// 2.再去抓Header的欄位名稱
				String firstRowName = HeaderName.get(key);
			
				
				

//				// 1.先抓現在第幾個Column
//				int cellNum = cell.getColumnIndex();
//				// 2.再去抓Header的欄位名稱
//				String firstRowName = HeaderName.get(key);
//
//				// 3.判斷是否為需要抓的欄位
//				if (!queryColArray.contains(firstRowName) || firstRowName == null) {
//					continue;
//				}
//			cell=row.getCell(cellNum,MissingCellPolicy.CREATE_NULL_AS_BLANK);

//				System.out.println(key + " : " + HeaderName.get(key));

//				String firstRowName1 = key.toString();
//				if (!queryColArray.contains(HeaderName.get(key)) || HeaderName.get(key) == null) {
//					continue;

//			System.out.println(HeaderName.get(key));

				// "TWS AD Name,JCL Name,STEP Name,PROGRAM Name,DISP Status"
				// 抓到的欄位如果是JCL Name 會需要做空值塞值
//			if (firstRowName.equals("JCL")) {
//				if (cell.getStringCellValue().isEmpty() || cell.getStringCellValue() == null) {
//					cell.setCellValue(JCLNameLast);
//				} else {
//					JCLNameLast = cell.getStringCellValue();
//				}
//			}
				
				
				
				// 如果是DISP Status，要抓DSN的值帶過來
//			if (firstRowName.equals("DISP_Initial")) {
//				switch (firstRowName) {
//				case "MOD":
//					firstRowName = "OUTPUT FILE";
//					cell.setCellValue(row.getCell(dnsIndex).getStringCellValue());
//					break;
//				case "OLD":
//					firstRowName = "INPUT FILE";
//					cell.setCellValue(row.getCell(dnsIndex).getStringCellValue());
//					break;
//				case "SHR":
//					firstRowName = "INPUT FILE";
//					cell.setCellValue(row.getCell(dnsIndex).getStringCellValue());
//					break;
//				case "TLB645":
//					break;
//				default:
//					break;
//				}
//			}
//			if (HeaderName.get(key).equals("IMS_Get")) {
//			System.out.println("1111");
//			System.out.println(HeaderName.get(21));
//				if(cell.getStringCellValue()==null) {
//				cell.setCellValue("123");
//				System.out.println(cell.getStringCellValue());
//			}
//			}
				//找內容  只要針對欄位 ( Delete_Reason_EXEC_DD)為空白的，再產生word資料即可
//				if (HeaderName.get(key).equals("Delete_Reason_EXEC_DD")) {
//					
//					if(!cell.getStringCellValue().equals("")) {
//						excelDateMap.remove(excelDateMap);
//						}
//						
//					}

					
				excelDateMap.put(firstRowName, cell.getStringCellValue());
			}	
		}
	
//		System.out.println(excelDateMap);
//		//找內容  只要針對欄位 ( Delete_Reason_EXEC_DD)為空白的，再產生word資料即可
//		if (HeaderName.get(key).equals("Delete_Reason_EXEC_DD")) {
//			
//			if(!cell.getStringCellValue().equals("")) {
//				excelDateMap.remove(excelDateMap);
//				}
//				
//			}
		
		return excelDateMap;
	}

//	D:\si1204\Desktop\單元測試個案\Batch
	public void createWord(List<Map<String, String>> excelList) {
		if (!new File(destFileFolderPath + "/" + excelList.get(0).get("Sprint_No") + "_"
				+ excelList.get(0).get("測試案例_L2") + "/" + excelList.get(0).get("AD")).exists()) {
			new File(destFileFolderPath + "/" + excelList.get(0).get("Sprint_No") + "_"
					+ excelList.get(0).get("測試案例_L2") + "/" + excelList.get(0).get("AD")).mkdirs();
		}
		try (InputStream is = new FileInputStream(tempFileFolderPath);
				OutputStream os = new FileOutputStream(destFileFolderPath + "/" + excelList.get(0).get("Sprint_No")
						+ "_" + excelList.get(0).get("測試案例_L2") + "/" + excelList.get(0).get("AD") + "/"
						+ excelList.get(0).get("AD") + "_" + excelList.get(0).get("JCL") + ".docx");) {
			XWPFDocument doc = XWPFUtils.openDoc(is);
			List<XWPFParagraph> xwpfParas = doc.getParagraphs();

			List<Map<String, String>> Catalog = excelList.stream()
					.filter(item -> item.get("DISP").equals("SHR") || item.get("DISP").equals("OLD"))
					.collect(Collectors.toList());
			List<Map<String, String>> inputList = excelList
					.stream().filter(item -> item.get("Open_Mode").equals("INPUT")
							|| item.get("Open_Mode").equals("I-O") || item.get("Open_Mode").equals("INPUT,OUTPUT"))
					.collect(Collectors.toList());

			List<Map<String, String>> ouputList = excelList.stream()
					.filter(item -> item.get("Open_Mode").equals("OUTPUT") || item.get("Open_Mode").equals("I-O")
							|| item.get("Open_Mode").equals("INPUT,OUTPUT")
							|| item.get("Open_Mode").equals("I-O,INPUT"))
					.collect(Collectors.toList());
//			List<Map<String, String>> imdbsTable = excelList.stream()
//					.filter(item -> !item.get("IMS_Get").equals("") || !item.get("IMS_Insert").equals("")
//							|| !item.get("IMS_Update").equals("") || !item.get("IMS_Delete").equals(""))
//					.distinct()
//					.collect(Collectors.toList());
//
//			List<Map<String, String>> db2Table = excelList.stream()
//					.filter(item -> !item.get("DB2_Include").equals("") || !item.get("DB2_Select").equals("")
//							|| !item.get("DB2_Insert").equals("") || !item.get("DB2_Update").equals("")
//							|| !item.get("DB2_Delete").equals(""))
//					.distinct().collect(Collectors.toList());
//			List<Map<String, String>> test = excelList.stream().collect(Collectors.toList());
//			System.out.println(test.toString());
			
			// !"".equals(item.get("Error_code"))反過來寫才不會錯誤
			List<Map<String, String>> errorCode = excelList.stream()
					.filter(item->(!"".equals(item.get("Error_code")) && item.get("Error_code") != null) 
							|| (!"".equals(item.get("異常處理")) && item.get("異常處理")!=null)
							|| (!"".equals(item.get("異常處理方式")) && item.get("異常處理方式")!=null)
							|| (!"".equals(item.get("通知方式")) && item.get("通知方式")!=null)
							|| (!"".equals(item.get("通知處理方式")) && item.get("通知處理方式")!=null)
							|| (!"".equals(item.get("Mobile_Number_1st")) && item.get("Mobile_Number_1st")!=null)
							|| (!"".equals(item.get("Mobile_Number_2nd")) && item.get("Mobile_Number_2nd")!=null))
					.filter(Objects::nonNull)
					
					.collect(Collectors.toList());
			
			// !item.get("Error_code").equals("") filter不能這樣寫,會nullpointerexception
//			List<Map<String, String>> errorCode = excelList.stream()
//					.filter(item->(!item.get("Error_code").equals("") && item.get("Error_code") != null) 
//							|| (!item.get("異常處理").equals("") && item.get("異常處理")!=null)
//							|| (!item.get("異常處理方式").equals("") && item.get("異常處理方式")!=null)
//							|| (!item.get("通知方式").equals("") && item.get("通知方式")!=null)
//							|| (!item.get("通知處理方式").equals("")&& item.get("通知處理方式")!=null)
//							|| (!item.get("Mobile_Number_1st").equals("") && item.get("Mobile_Number_1st")!=null)
//							|| (!item.get("Mobile_Number_2nd").equals("") && item.get("Mobile_Number_2nd")!=null))
//					.filter(Objects::nonNull)
//					.collect(Collectors.toList());
			
			
			
			
		
//			System.out.println(errorCode.toString());
			for (XWPFParagraph xwpfParagraph : xwpfParas) {
				String itemText = xwpfParagraph.getText();
				switch (itemText) {
				case "${catalogTable}":
					XWPFUtils.replaceTable(doc, itemText, Catalog, tableOutputKey);
					break;

				case "${dataTable}":
					XWPFUtils.replaceTable(doc, itemText, inputList, tableOutputKey);
					break;

				case "${resultTable}":
					XWPFUtils.replaceTable(doc, itemText, ouputList, table3OutputKey);
					break;
//				case "${imdbsTable}":
//					XWPFUtils.replaceTable(doc, itemText, imdbsTable, table4OutputKey);
//					break;
//				case "${db2Table}":
//					XWPFUtils.replaceTable(doc, itemText, db2Table, table5OutputKey);
//					break;
				case "${Error_code}":
					XWPFUtils.replaceTable(doc, itemText, errorCode, table6OutputKey);
					break;	
					
				}
			}

			Map<String, Object> data = new HashMap<>();

			data.put("${測試案例_L2}", excelList.get(0).get("測試案例_L2"));
			data.put("${系統別}", excelList.get(0).get("系統別"));
			data.put("${AD_NAME}", excelList.get(0).get("AD"));

			if (StringUtils.isEmpty(excelList.get(0).get("AD Description").trim())) {
				data.put("${AD_Description}", " ");
			} else {
				data.put("${AD_Description}", "<<" + excelList.get(0).get("AD Description") + ">>");
			}
			data.put("${JCL_NAME}", excelList.get(0).get("JCL"));

			if (StringUtils.isEmpty(excelList.get(0).get("JCL Description").trim())) {
				data.put("${JCL_Description}", " ");
			} else {
				data.put("${JCL_Description}", "<<" + excelList.get(0).get("JCL Description") + ">>");

			}
//			data.put("${IMS_Get}", excelList.get(0).get("IMS_Get"));
//			data.put("${DB2_Include}", excelList.get(0).get("DB2_Include"));
//			System.out.println("111111111111111"+excelList.get(0).get("DB2_Include"));
//			data.put("${DB2_Select}", excelList.get(0).get("DB2_Select"));
//			System.out.println("222222222222222"+excelList.get(0).get("IMS_Get"));
			
//			System.out.println("Test Map Data:"+data.toString());
			// 取代資料
			XWPFUtils.replaceInPara(doc, data);

			
            
            Map<String, Object> data1 = new HashMap<>();
            data1.put("${IMS_Get}", excelList.get(0).get("IMS_Get"));
            data1.put("${IMS_Insert}", excelList.get(0).get("IMS_Insert"));
            data1.put("${IMS_Update}", excelList.get(0).get("IMS_Update"));
            data1.put("${IMS_Delete}", excelList.get(0).get("IMS_Delete"));
            data1.put("${DB2_Include}", excelList.get(0).get("DB2_Include"));
            data1.put("${DB2_Select}", excelList.get(0).get("DB2_Select"));
            data1.put("${DB2_Insert}", excelList.get(0).get("DB2_Insert"));
            data1.put("${DB2_Update}", excelList.get(0).get("DB2_Update"));
            data1.put("${DB2_Delete}", excelList.get(0).get("DB2_Delete"));
            XWPFUtils.searchAndReplace(doc, data1);
          

          
			doc.write(os);
		} catch (FileNotFoundException e) {
			logger.info(e.toString());
			e.printStackTrace();
		} catch (IOException e) {
			logger.info(e.toString());
			e.printStackTrace();
		}
	}

}
