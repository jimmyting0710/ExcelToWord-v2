package com.example.utils;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.ooxml.POIXMLDocument;
import org.apache.poi.xwpf.usermodel.BodyElementType;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.TableWidthType;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlCursor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHeight;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTbl;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblGrid;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblGridCol;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTrPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTVMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STVerticalJc;

public class XWPFUtils {

	/**
	 * 打开word文档
	 * 
	 * @param path 文档所在路径
	 * @return
	 * @throws IOException
	 * @Author Huangxiaocong 2018年12月1日 下午12:30:07
	 */
	public XWPFDocument openDoc(InputStream is) throws IOException {
		return new XWPFDocument(is);
	}

	/**
	 * 替換段落裡面的變數
	 *
	 * @param doc    要替換的文件
	 * @param params 引數
	 */
	public static void replaceInPara(XWPFDocument doc, Map<String, Object> params) {
		Iterator<XWPFParagraph> iterator = doc.getParagraphsIterator();
		XWPFParagraph para;
		while (iterator.hasNext()) {
			para = iterator.next();
			replaceInPara(para, params);
		}
	}

	/**
	 * 替換段落裡面的變數
	 *
	 * @param para   要替換的段落
	 * @param params 引數
	 */
	private static void replaceInPara(XWPFParagraph para, Map<String, Object> params) {
		Matcher matcher = matcher(para.getParagraphText());
		String runText = "";
		String keyString;
		if (matcher.find()) {
			runText = getSpliString(para);
			while ((matcher = matcher(runText)).find()) {
				keyString = matcher.group(0);
				runText = matcher.replaceFirst(Matcher.quoteReplacement(String.valueOf(params.get(keyString))));
			}
			// 直接呼叫XWPFRun的setText()方法設定文字時，在底層會重新建立一個XWPFRun，把文字附加在當前文字後面，
			// 所以我們不能直接設值，需要先刪除當前run,然後再自己手動插入一個新的run。
			para.insertNewRun(0).setText(runText);
		}

	}

	public void replaceTable(XWPFDocument doc, String tagString, List<Map<String, String>> dataList,
			List<String> queryColArray) {
		List<XWPFParagraph> paras = doc.getParagraphs();
		for (XWPFParagraph para : paras) {
			//
			String runString = para.getParagraphText();
			// Match Paragraph Test
			Matcher matcher = matcher(para.getParagraphText());
			while (matcher.find()) {
				runString = matcher.group(0);
			}
			List<XWPFRun> runs = para.getRuns();
//				String runString = run.getText(0).trim();
			if (runString != null) {
				if (runString.indexOf(tagString) >= 0) {
					for (XWPFRun run : runs) {
						run.setText(runString.replace(tagString, ""), 0);
					}
					XmlCursor cursor = para.getCTP().newCursor();
					XWPFTable table = doc.insertNewTbl(cursor);
					
					//自設表格寬度，不設定會auto
//					CTTblPr tablePr = table.getCTTbl().getTblPr();
//					CTTblWidth width = tablePr.addNewTblW();
//					width.setW(BigInteger.valueOf(8500));
					fillHeaderData(table, dataList, queryColArray);
					fillTableData(table, dataList, queryColArray);
					setTableLocation(table, "center");
				}
			}
		}
	}

	public static String getSpliString(XWPFParagraph para) {
		List<XWPFRun> runs;
		runs = para.getRuns();
		StringBuilder runText = new StringBuilder();
		if (runs.size() > 0) {
			int j = runs.size();
			for (int i = 0; i < j; i++) {
				XWPFRun run = runs.get(0);
				String i1 = run.toString();
				runText.append(i1);
				para.removeRun(0);

			}

		}
		return runText.toString();
	}

	/**
	 * 正則匹配字串
	 *
	 * @param str
	 * @return
	 */
	public static Matcher matcher(String str) {
		Pattern pattern = Pattern.compile("\\$\\{(.+?)\\}", Pattern.CASE_INSENSITIVE);
		Matcher matcher = pattern.matcher(str);
		return matcher;
	}

	/**
	 * 刪除指定位置的表格，被刪除表格後的索引位置
	 * 
	 * @param document
	 * @param pos
	 * @Author Huangxiaocong 2018年12月1日 下午10:32:43
	 */
//	public void deleteTableByIndex(XWPFDocument document, int pos) {
//		Iterator<IBodyElement> bodyElement = document.getBodyElementsIterator();
//		int eIndex = 0, tableIndex = -1;
//		while (bodyElement.hasNext()) {
//			IBodyElement element = bodyElement.next();
//			BodyElementType elementType = element.getElementType();
//			if (elementType == BodyElementType.TABLE) {
//				tableIndex++;
//				if (tableIndex == pos) {
//					break;
//				}
//			}
//			eIndex++;
//		}
//		document.removeBodyElement(eIndex);
//	}

	/**
	 * 獲得指定位置的表格
	 * 
	 * @param document
	 * @param index
	 * @return
	 * @Author Huangxiaocong 2018年12月1日 下午10:34:14
	 */
//	public XWPFTable getTableByIndex(XWPFDocument document, int index) {
//		List<XWPFTable> tableList = document.getTables();
//		if (tableList == null || index < 0 || index > tableList.size()) {
//			return null;
//		}
//		return tableList.get(index);
//	}

	/**
	 * 得到表格的內容（第一次跨行單元格視為一個，第二次跳過跨行合併的單元格）
	 * 
	 * @param table
	 * @return
	 * @Author Huangxiaocong 2018年12月1日 下午10:46:41
	 */
//	public List<List<String>> getTableRConten(XWPFTable table) {
//		List<List<String>> tableContextList = new ArrayList<List<String>>();
//		for (int rowIndex = 0, rowLen = table.getNumberOfRows(); rowIndex < rowLen; rowIndex++) {
//			XWPFTableRow row = table.getRow(rowIndex);
//			List<String> cellContentList = new ArrayList<String>();
//			for (int colIndex = 0, colLen = row.getTableCells().size(); colIndex < colLen; colIndex++) {
//				XWPFTableCell cell = row.getCell(colIndex);
//				CTTc ctTc = cell.getCTTc();
//				if (ctTc.isSetTcPr()) {
//					CTTcPr tcPr = ctTc.getTcPr();
//					if (tcPr.isSetHMerge()) {
//						CTHMerge hMerge = tcPr.getHMerge();
//						if (STMerge.RESTART.equals(hMerge.getVal())) {
//							cellContentList.add(getTableCellContent(cell));
//						}
//					} else if (tcPr.isSetVMerge()) {
//						CTVMerge vMerge = tcPr.getVMerge();
//						if (STMerge.RESTART.equals(vMerge.getVal())) {
//							cellContentList.add(getTableCellContent(cell));
//						}
//					} else {
//						cellContentList.add(getTableCellContent(cell));
//					}
//				}
//			}
//			tableContextList.add(cellContentList);
//		}
//		return tableContextList;
//	}

	/**
	 * 獲得一個表格的單元格的內容
	 * 
	 * @param cell
	 * @return
	 * @Author Huangxiaocong 2018年12月2日 下午7:39:23
	 */
//	public String getTableCellContent(XWPFTableCell cell) {
//		StringBuffer sb = new StringBuffer();
//		List<XWPFParagraph> cellParagList = cell.getParagraphs();
//		if (cellParagList != null && cellParagList.size() > 0) {
//			for (XWPFParagraph xwpfPr : cellParagList) {
//				List<XWPFRun> runs = xwpfPr.getRuns();
//				if (runs != null && runs.size() > 0) {
//					for (XWPFRun xwpfRun : runs) {
//						sb.append(xwpfRun.getText(0));
//					}
//				}
//			}
//		}
//		return sb.toString();
//	}

	/**
	 * 得到表格內容，合併後的單元格視為一個單元格
	 * 
	 * @param table
	 * @return
	 * @Author Huangxiaocong 2018年12月2日 下午7:47:19
	 */
//	public List<List<String>> getTableContent(XWPFTable table) {
//		List<List<String>> tableContentList = new ArrayList<List<String>>();
//		for (int rowIndex = 0, rowLen = table.getNumberOfRows(); rowIndex < rowLen; rowIndex++) {
//			XWPFTableRow row = table.getRow(rowIndex);
//			List<String> cellContentList = new ArrayList<String>();
//			for (int colIndex = 0, colLen = row.getTableCells().size(); colIndex < colLen; colIndex++) {
//				XWPFTableCell cell = row.getCell(colIndex);
//				cellContentList.add(getTableCellContent(cell));
//			}
//			tableContentList.add(cellContentList);
//		}
//		return tableContentList;
//	}

	/**
	 * 跨列合併
	 * 
	 * @param table
	 * @param row      所合併的行
	 * @param fromCell 起始列
	 * @param toCell   終止列
	 * @Description
	 * @Author Huangxiaocong 2018年11月26日 下午9:23:22
	 */
//	public void mergeCellsHorizontal(XWPFTable table, int row, int fromCell, int toCell) {
//		for (int cellIndex = fromCell; cellIndex <= toCell; cellIndex++) {
//			XWPFTableCell cell = table.getRow(row).getCell(cellIndex);
//			if (cellIndex == fromCell) {
//				cell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.RESTART);
//			} else {
//				cell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.CONTINUE);
//			}
//		}
//	}

	/**
	 * 跨行合併
	 * 
	 * @param table
	 * @param col     合併的列
	 * @param fromRow 起始行
	 * @param toRow   終止行
	 * @Description
	 * @Author Huangxiaocong 2018年11月26日 下午9:09:19
	 */
//	public void mergeCellsVertically(XWPFTable table, int col, int fromRow, int toRow) {
//		for (int rowIndex = fromRow; rowIndex <= toRow; rowIndex++) {
//			XWPFTableCell cell = table.getRow(rowIndex).getCell(col);
//			// 第一個合併單元格用重啟合併值設定
//			if (rowIndex == fromRow) {
//				cell.getCTTc().addNewTcPr().addNewVMerge().setVal(STMerge.RESTART);
//			} else {
//				// 合併第一個單元格的單元被設定為“繼續”
//				cell.getCTTc().addNewTcPr().addNewVMerge().setVal(STMerge.CONTINUE);
//			}
//		}
//	}

	/**
	 * @Description: 建立表格,建立後表格至少有1行1列,設定列寬
	 */
//	public XWPFTable createTable(XWPFDocument xdoc, int rowSize, int cellSize, boolean isSetColWidth, int[] colWidths) {
//		XWPFTable table = xdoc.createTable(rowSize, cellSize);
//		if (isSetColWidth) {
//			CTTbl ttbl = table.getCTTbl();
//			CTTblGrid tblGrid = ttbl.addNewTblGrid();
//			for (int j = 0, len = Math.min(cellSize, colWidths.length); j < len; j++) {
//				CTTblGridCol gridCol = tblGrid.addNewGridCol();
//				gridCol.setW(new BigInteger(String.valueOf(colWidths[j])));
//			}
//		}
//		return table;
//	}

	/**
	 * @Description: 設定表格總寬度與水平對齊方式
	 */
//	public void setTableWidthAndHAlign(XWPFTable table, String width, STJc.Enum enumValue) {
//		CTTblPr tblPr = getTableCTTblPr(table);
//		// 表格寬度
//		CTTblWidth tblWidth = tblPr.isSetTblW() ? tblPr.getTblW() : tblPr.addNewTblW();
//		if (enumValue != null) {
//			CTJc cTJc = tblPr.addNewJc();
//			cTJc.setVal(enumValue);
//		}
//		// 設定寬度
//		tblWidth.setW(new BigInteger(width));
//		tblWidth.setType(STTblWidth.DXA);
//	}

	/**
	 * @Description: 得到Table的CTTblPr,不存在則新建
	 */
//	public CTTblPr getTableCTTblPr(XWPFTable table) {
//		CTTbl ttbl = table.getCTTbl();
//		// 表格屬性
//		CTTblPr tblPr = ttbl.getTblPr() == null ? ttbl.addNewTblPr() : ttbl.getTblPr();
//		return tblPr;
//	}

	/**
	 * 設定表格行高
	 * 
	 * @param infoTable
	 * @param heigth    高度
	 * @param vertical  表格內容的顯示方式：居中、靠右...
	 * @Author Huangxiaocong 2018年12月16日
	 */
//	public void setTableHeight(XWPFTable infoTable, int heigth, STVerticalJc.Enum vertical) {
//		List<XWPFTableRow> rows = infoTable.getRows();
//		for (XWPFTableRow row : rows) {
//			CTTrPr trPr = row.getCtRow().addNewTrPr();
//			CTHeight ht = trPr.addNewTrHeight();
//			ht.setVal(BigInteger.valueOf(heigth));
//			List<XWPFTableCell> cells = row.getTableCells();
//			for (XWPFTableCell tableCell : cells) {
//				CTTcPr cttcpr = tableCell.getCTTc().addNewTcPr();
//				cttcpr.addNewVAlign().setVal(vertical);
//			}
//		}
//	}

	public void fillHeaderData(XWPFTable table, List<Map<String, String>> tableData, List<String> HeaderName) {
		XWPFTableRow headerRow = table.getRow(0);
		XWPFTableCell cell;
		for (int i = 0; i < HeaderName.size(); i++) {
//			if(HeaderName.get(i).equals("DSN")) {   //不印dsn
//				continue;
//			}
			if (headerRow.getCell(i) == null) {
				cell = headerRow.createCell();
			} else {
				cell = headerRow.getCell(i);
			}
			XWPFParagraph cellParagraph = cell.getParagraphArray(0);
			XWPFRun cellParagraphRun = cellParagraph.createRun();
			cellParagraphRun.setText(HeaderName.get(i));

		}
	}

	/**
	 * 往表格中填充数据
	 * 
	 * @param table
	 * @param tableData
	 * @Author Huangxiaocong 2018年12月16日
	 */
	public void fillTableData(XWPFTable table, List<Map<String, String>> tableData, List<String> HeaderName) {
		for (int i = 0; i < tableData.size(); i++) {
			XWPFTableRow row = table.createRow();
			Map<String, String> item = tableData.get(i);
			for (int j = 0; j < row.getTableCells().size(); j++) {
				XWPFTableCell cell = row.getCell(j);
				XWPFParagraph cellParagraph = cell.getParagraphArray(0);
				XWPFRun cellParagraphRun = cellParagraph.createRun();
				cellParagraphRun.setText(item.get(HeaderName.get(j)));
			}
		}
	}

	/**
	 * 設定表格位置
	 *
	 * @param xwpfTable
	 * @param location  整個表格居中center,left居左，right居右，both兩端對齊
	 */
	public static void setTableLocation(XWPFTable xwpfTable, String location) {
		CTTbl cttbl = xwpfTable.getCTTbl();
		CTTblPr tblpr = cttbl.getTblPr() == null ? cttbl.addNewTblPr() : cttbl.getTblPr();
		CTJc cTJc = tblpr.addNewJc();
		cTJc.setVal(STJc.Enum.forString(location));
	}

//	public void copy(XWPFTable table, XWPFTableRow sourceRow, int rowIndex) {
//		// 在表格指定位置新增一行
//		XWPFTableRow targetRow = table.insertNewTableRow(rowIndex);
//		// 复制行属性
//		targetRow.getCtRow().setTrPr(sourceRow.getCtRow().getTrPr());
//		List<XWPFTableCell> cellList = sourceRow.getTableCells();
//		if (null == cellList) {
//			return;
//		}
//		// 复制列及其属性和内容
//		XWPFTableCell targetCell = null;
//		for (XWPFTableCell sourceCell : cellList) {
//			targetCell = targetRow.addNewTableCell();
//			// 列属性
//			targetCell.getCTTc().setTcPr(sourceCell.getCTTc().getTcPr());
//			// 段落属性
//			if (sourceCell.getParagraphs() != null && sourceCell.getParagraphs().size() > 0) {
//				targetCell.getParagraphs().get(0).getCTP().setPPr(sourceCell.getParagraphs().get(0).getCTP().getPPr());
//				if (sourceCell.getParagraphs().get(0).getRuns() != null
//						&& sourceCell.getParagraphs().get(0).getRuns().size() > 0) {
//					XWPFRun cellR = targetCell.getParagraphs().get(0).createRun();
//					cellR.setText(sourceCell.getText());
//					cellR.setBold(sourceCell.getParagraphs().get(0).getRuns().get(0).isBold());
//				} else {
//					targetCell.setText(sourceCell.getText());
//				}
//			} else {
//				targetCell.setText(sourceCell.getText());
//			}
//		}
//
//	}

	
	
	//自已後來找的
	 public static void searchAndReplace(XWPFDocument document,Map<String, String> data1) {
		try {

			/**
			 * 替換表格中的指定文字
			 */
			Iterator<XWPFTable> itTable = document.getTablesIterator(); //獲得Word的表格
			
			while (itTable.hasNext()) { //遍歷表格
				XWPFTable table = (XWPFTable) itTable.next();
				int count = table.getNumberOfRows(); //獲得表格總行數
				for (int i = 0; i < count; i++) {    //遍歷表格的每一行
					XWPFTableRow row = table.getRow(i);  //獲得表格的行
					List<XWPFTableCell> cells = row.getTableCells();  //在行元素中，獲得表格的單元格
					for (XWPFTableCell cell : cells) {        //遍歷單元格
						for (Entry<String, String> e : data1.entrySet()) {
							if (cell.getText().equals(e.getKey())) {   //如果單元格中的變數和‘鍵’相等，就用‘鍵’所對應的‘值’代替。
								cell.removeParagraph(0);      //所以這裡就要求每一個單元格只能有唯一的變數。
								cell.setText((String) e.getValue());
							}
						}
					}
				}
			}
//			FileOutputStream outStream = null;
//            outStream = new FileOutputStream(destPath);
//            document.write(outStream);
//            outStream.close();
		} catch (Exception e) {
			e.printStackTrace();
		}

	}
}