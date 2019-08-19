package com.wx.service;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collection;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;

import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

import com.mks.api.response.APIException;
import com.wx.ui.TestCaseImport;
import com.wx.util.ExceptionUtil;
import com.wx.util.MKSCommand;


public class ExcelUtil {

	private static List<String> caseFields = new ArrayList<>();
	private static List<String> stepFields = new ArrayList<>();
	private static List<String> resultFields = new ArrayList<>();
	private static List<String> allHeaders = new ArrayList<>();
	public String[][] tableFields = null;
	public static Map<String,String> hasParentField = new HashMap<String,String>();
	private static Map<String, Map<String, String>> headerConfig = new HashMap<>();
	private static final String FIELD_CONFIG_FILE = "FieldMapping.xml";
	private static final String CATEGORY_CONFIG_FILE = "Category.xml";
	private static final String TEST_STEP = "Test Step";
	private static final String TEST_RESULT = "Test Result";
	private static final String TEST_SESSION = "Test Session";
	private static final String SESSIONN_STATE = "In Testing";
	private static final String PARENT_FIELD = "parentField";
	private static final String NEED_FIELD_SET = "needFieldSet";
	private static final String TEST_CASE_ID = "ID";
	private static final String TEXT = "Text";
	private static final String VERDICT = "Verdict";
	private static final String ExpectedResults = "Expected Results";
	private static final String SESSION_ID = "Session ID";
	private static final String SEQUENCE = "Sequence";
	
	private static final String SEQUENCE_FIELD = "Customer Sequence";
	private static final String CHILD = "Child";
	
	private static final List<String> CURRENT_CATEGORIES = new ArrayList<String>();//记录导入对象的正确Category
	
	private static final Map<String,List<String>> PICK_FIELD_RECORD = new HashMap<String,List<String>>();
	
	private static final Map<String,String> FIELD_TYPE_RECORD = new HashMap<String,String>();
	
	public static final List<String> RICH_FIELDS = new ArrayList<String>();
	
	private static String IMPORT_DOC_TYPE = "Test Suite";
	
	private static String CONTENT_TYPE ;
	
	private static final String SPERATOR = "_";
	
	private static final String INIT_CONTENT_STATE = "New";
	
	private static final String INIT_STEP_STATE = "Active";
	
	private Map<String,CellRangeAddress> cellRangeMap = new HashMap<String,CellRangeAddress>();
	
	private static final List<String> USER_FULLNAME_RECORD = new ArrayList<String>();
	private static  boolean IS_USER = false ;
	private static  boolean RELATIONSHIP_MISTAKEN = false ;
	public static final Logger logger = Logger.getLogger(ExcelUtil.class);
	
	/**
	 * 利用Jsoup解析配置文件，得到相应的参数，为Type选项和创建Document提供信息 (1)
	 * Document:Type,Project,State,Shared Category (2) Content:Type 负责人：汪巍
	 * 
	 * @return
	 * @throws Exception 
	 */
	public List<String> parsFieldMapping() throws Exception {
		
		ExcelUtil.logger.info("start to parse xml : " + FIELD_CONFIG_FILE);
		Document doc = DocumentBuilderFactory.newInstance().newDocumentBuilder()
				.parse(ExcelUtil.class.getClassLoader().getResourceAsStream( FIELD_CONFIG_FILE));
		Element root = doc.getDocumentElement();
		List<String> typeList = new ArrayList<String>();
		if (root == null)
			return typeList;
		// 得到xml配置
		NodeList importTypes = root.getElementsByTagName("importType");  // 拿到mapping里面所有的 ImportType
		if (importTypes == null || importTypes.getLength() == 0) {
			throw new Exception("Can't not parse xml because of don't has \"importType\"");
		}else {
			// 循环 刚才拿到的所有ImportType
			for (int j = 0; j < importTypes.getLength(); j++) {
				Element importType = (Element) importTypes.item(j);
				// 获取XML 文件的name 和  Type
				String documentType = importType.getAttribute("type");
				IMPORT_DOC_TYPE = documentType;
				NodeList excelFields = importType.getElementsByTagName("excelField");
				try {
					if (excelFields == null || excelFields.getLength() == 0) {
						throw new Exception("Can't not parse xml because of don't has \"excelField\"");
					} else {
						tableFields = new String[excelFields.getLength()][2];
						for (int i = 0; i < excelFields.getLength(); i++) {
							Element fields = (Element) excelFields.item(i);
							String name = fields.getAttribute("name");
							Map<String, String> map = new HashMap<>();
							String type = fields.getAttribute("type");
							allHeaders.add(name);
							map.put("type", type);
							if (TEST_STEP.equals(type) && !stepFields.contains(name)) {
								stepFields.add(name);
							} else if (TEST_RESULT.equals(type) && !resultFields.contains(name)) {
								resultFields.add(name);
							}else if (!TEST_STEP.equals(type) && !TEST_RESULT.equals(type) && !caseFields.contains(name)) {
								caseFields.add(name);
								CONTENT_TYPE = type;
							}
							String field = fields.getAttribute("field");
							map.put("field", field);
							// 获取 excelField 的  onlyCreate 属性 ， 若没有填写则默认为 false 
							String onlyCreate = fields.getAttribute("onlyCreate");
							if(onlyCreate == null || onlyCreate.equals("") ) {
								map.put("onlyCreate", "false");
							}else {
								map.put("onlyCreate", onlyCreate);
							}
							String overRide = fields.getAttribute("overRide");
							if(overRide == null || overRide.equals("") ) {
								map.put("overRide", "true");
							}else {
								map.put("overRide", overRide);
							}
							tableFields[i][0] = name;
							tableFields[i][1] = field;
							if(fields.hasAttribute(PARENT_FIELD)) {
								String parentField = fields.getAttribute(PARENT_FIELD);
								map.put(PARENT_FIELD, parentField);
								tableFields[i][1] = parentField;
								hasParentField.put(field, parentField);
							}
							if(fields.hasAttribute(NEED_FIELD_SET)) {
								String needField = fields.getAttribute(NEED_FIELD_SET);
								map.put(NEED_FIELD_SET, needField);
							}
							headerConfig.put(name, map);
						}
					}
				} catch (ParserConfigurationException e) {
					logger.error("parse config file exception", e);
				} catch (SAXException e) {
					logger.error("get config file exception", e);
				} catch (IOException e) {
					logger.error("io exception", e);
				} finally {
					logger.info("get info : \nheaderConfig : " + headerConfig);
				}
			}
		}
		return typeList;
	}
	
	/**
	 * Description 查询当前要导入类型的 正确Category
	 * @param documentType
	 * @throws Exception
	 */
	public void parseCurrentCategories(String documentType) throws Exception{
		Document doc = DocumentBuilderFactory.newInstance().newDocumentBuilder()
				.parse(ExcelUtil.class.getClassLoader().getResourceAsStream(CATEGORY_CONFIG_FILE));
		Element root = doc.getDocumentElement();
		// 得到xml配置
		NodeList importTypes = root.getElementsByTagName("documentType");
		for (int j = 0; j < importTypes.getLength(); j++) {
			Element importType = (Element) importTypes.item(j);
			String typeName = importType.getAttribute("name");
			if(typeName.equals(documentType)){
				NodeList categoryNodes = importType.getElementsByTagName("category");
				for (int i = 0; i < categoryNodes.getLength(); i++) {
					Element categoryNode = (Element) categoryNodes.item(i);
					CURRENT_CATEGORIES.add(categoryNode.getAttribute("name"));
				}
			}
		}	
	}
	
	/**
	 * 获得Excel中的数据
	 * 
	 * @param filePath
	 * @return
	 * @throws BiffException
	 * @throws IOException
	 */
	public List<List<Map<String, Object>>> parseExcel(File file) throws Exception {
		Workbook wb = null;
		String fileName = file.getName();
		if(fileName.endsWith(".xlsx")) {
			wb = new XSSFWorkbook(file);
		}else if(fileName.endsWith(".xls")) {
			wb = new HSSFWorkbook(new FileInputStream(file));
		}
		List<List<Map<String, Object>>> datas = new ArrayList<>();
		List<Map<String, Object>> list = null;
		Iterator<Sheet> iterator = wb.sheetIterator();
		Integer IDIndex = 0;
		if(allHeaders.indexOf(TEST_CASE_ID)>-1)
			IDIndex = allHeaders.indexOf(TEST_CASE_ID);
		while(iterator.hasNext()){
			list = new ArrayList<>();
			Sheet sheet = iterator.next();
			String sheetName = sheet.getSheetName();
			Map<String,Object> headerMap = new HashMap<String,Object>();
			headerMap.put(SEQUENCE_FIELD, sheetName);
			headerMap.put("Text", sheetName);
			headerMap.put("Category", "Heading");
			list.add(headerMap);
			List<CellRangeAddress> mergeList = sheet.getMergedRegions();
			cellRangeMap = new HashMap<String,CellRangeAddress>();
			int rowNum = this.getRealRowNum(sheet,mergeList);
			int colNum = this.getRealColNum(sheet);
			int row = 1;
			Row firstRow = sheet.getRow(0);
			Row secondRow = sheet.getRow(1);
			int merge = getMergeRow(mergeList);
			if(merge > 0) {
				row = row + merge;
			}
			int col = 0;
			int endRow = row + rowNum;
			for ( ; row < endRow; row++) {
				try{
					Map<String, Object> map = new HashMap<>();
					Map<String,String> stepMap = null;
					Map<String,String> resultMap = null;
					List<Map<String,String>> stepList = new ArrayList<Map<String,String>>();
					List<Map<String,String>> resultList = new ArrayList<Map<String,String>>();
					//Test Case可关联多个Test Step信息，通过多行关联
					//Test Case可关系多个Test Result信息，通过多列关联
					int caseMerge = 0;
					CellRangeAddress IDCellRange = cellRangeMap.get(row + SPERATOR + IDIndex);
					if(IDCellRange != null){
						int endMergeRow = IDCellRange.getLastRow();
						caseMerge = endMergeRow - row;
					}
					int temp = row ;
					for(; temp <= row + caseMerge; temp++){
						Row dataRow = sheet.getRow(temp);
						stepMap = new HashMap<String,String>();
						for ( col = 0; col < colNum; col++) {
							Cell fieldCell = firstRow.getCell(col);
							Cell secondCell = secondRow.getCell(col);
							String field = getCellVal(fieldCell);
							String secondFieldVal = getCellVal(secondCell);
							Cell valueCell = dataRow.getCell(col);
							String valueVal = getCellVal(valueCell);
							if("".equals(secondFieldVal) || "-".equals(secondFieldVal)){//这是Test Case数据
								Object value = map.get(field);
								if(value == null || "".equals(value)){
									map.put(field, valueVal);
								}else if(ExpectedResults.equals(field)){//如果是Expected Results，合并值
									String valueStr = (String) value;
									valueStr = valueStr + "\n" + valueVal;
								}
								
							}else if(secondFieldVal != null && stepFields.contains(secondFieldVal) ){
								if(!"".equals(valueVal))
									stepMap.put(secondFieldVal, valueVal);
							}else if(secondFieldVal != null && resultFields.contains(secondFieldVal)){//循环处理Test Result
								if(valueVal != null && !"".equals(valueVal)){
									if(SESSION_ID.equals(secondFieldVal)){
										resultMap = new HashMap<String,String>();
										resultList.add(resultMap);
									}
									resultMap.put(secondFieldVal, valueVal);
								}
							}
						}
						if(!stepMap.isEmpty())
							stepList.add(stepMap);
					}
					row = row + caseMerge;
					if(!stepList.isEmpty()) 
						map.put(TEST_STEP, stepList);
					if(!resultList.isEmpty()) 
						map.put(TEST_RESULT, resultList);
					list.add(map);
				}catch (Exception e){
					e.printStackTrace();
					System.out.println(row);
					System.out.println(col);
				}
			}
			datas.add(list);
		}
		return datas;
	}
	
	@SuppressWarnings("deprecation")
	public String getCellVal(Cell cell) {
		String value = "";
		if(cell != null){
			switch (cell.getCellType()) {
				case Cell.CELL_TYPE_STRING:
					value = cell.getStringCellValue();
					break;
				case Cell.CELL_TYPE_BLANK:
					break;
				case Cell.CELL_TYPE_FORMULA:
					value = String.valueOf(cell.getCellFormula());
					break;
				case Cell.CELL_TYPE_NUMERIC:
					value = String.valueOf(Math.round(cell.getNumericCellValue()));//当前项目 没有Number类型，只有String。取整
					break;
				case Cell.CELL_TYPE_BOOLEAN:
					value = String.valueOf(cell.getBooleanCellValue());
					break;
			}
		}
		return value;
	}
	
	/**
	 * 处理Excel中的数据，将Test Step信息和Test Case信息拆分开
	 * 
	 * @param data
	 * @return
	 */
	@SuppressWarnings("unchecked")
	public List<List<Map<String, Object>>> dealExcelData(List<List<Map<String, Object>>> datas) {
		List<List<Map<String, Object>>> newDatas = new ArrayList<>();
		List<Map<String, Object>> newData = null;
		if(!datas.isEmpty()){
			for(List<Map<String, Object>> data : datas){
				newData = new ArrayList<>();
				Map<String, Object> headerMap = data.get(0);
				data.remove(0);
				Map<String, Object> newMap = null;
				for(int i = 0; i < data.size(); i++ ) {
					Map<String,Object> rowMap = data.get(i);
					String caseID = (String)rowMap.get(TEST_CASE_ID);
					newMap = new HashMap<String,Object>();
					if(caseID != null && !"".equals(caseID)) {
						newMap.put("ID", caseID);
					}
					for (String header : caseFields) {
						Map<String, String> fieldConfig = headerConfig.get(header);
						if(fieldConfig != null) {
							String field = fieldConfig.get("field");
							String value = (String)rowMap.get(header);
							if (!"-".equals(field) && value != null && !"".equals(value)) {
								if(SEQUENCE.equals(header)){
									String subSequence = (String)rowMap.get("Sub-" + SEQUENCE);
									if(subSequence != null && !"".equals(subSequence)){
										value = value + SPERATOR + subSequence;
										String subSubSequence = (String)rowMap.get("Sub-Sub-" + SEQUENCE);
										if(subSubSequence != null && !"".equals(subSubSequence))
											value = value + SPERATOR + subSubSequence;
									}
								}
								newMap.put(field, value);
							}
						}
					}
					if(rowMap.containsKey(TEST_STEP)) {//Test Case包含有 Test Step信息
						Object steps = rowMap.get(TEST_STEP);
						if(steps instanceof List){
							List<Map<String, String>> currentSteps = (List<Map<String, String>> )steps;
							if(!currentSteps.isEmpty()) {//Test Case包含有 Test Step信息
								List<Map<String, String>> stepList = new ArrayList<Map<String, String>>();
								Map<String,String> stepMap = null;
								boolean hasVal = false;
								for(Map<String, String> map : currentSteps) {//循环处理Test Step信息
									hasVal = false;
									stepMap = new HashMap<String,String>();
									for(String header : stepFields) {
										if(map.containsKey(header)){
											Map<String, String> fieldConfig = headerConfig.get(header);
											if(fieldConfig != null) {
												String field = fieldConfig.get("field");
												String value = (String)map.get(header);
												if(value != null && !"".equals(value)) {
													stepMap.put(field, value);//存放非拼接字段
													hasVal = true;
												}
											}
										}
									}
									if(hasVal)
										stepList.add(stepMap);
								}
								newMap.put(TEST_STEP, stepList);
							}
						}
					}
					if(rowMap.containsKey(TEST_RESULT)) {//Test Case包含有 Test Result信息
						Object steps = rowMap.get(TEST_RESULT);
						if(steps instanceof List){
							List<Map<String, String>> currentResults = (List<Map<String, String>> )steps;
							if(!currentResults.isEmpty()) {//Test Case包含有 Test Result信息
								List<Map<String, String>> resultList = new ArrayList<Map<String, String>>();
								Map<String,String> resultMap = null;
								boolean hasVal = false;
								for(Map<String, String> map : currentResults) {//循环处理Test Result信息
									hasVal = false;
									resultMap = new HashMap<String,String>();
									for(String header : resultFields) {
										if(map.containsKey(header)){
											Map<String, String> fieldConfig = headerConfig.get(header);
											if(fieldConfig != null) {
												String field = fieldConfig.get("field");
												String value = (String)map.get(header);
												if(value != null && !"".equals(value)) {
													resultMap.put(field, value);//存放非拼接字段
													hasVal = true;
												}
											}
										}
									}
									if(hasVal)
										resultList.add(resultMap);
								}
								newMap.put(TEST_RESULT, resultList);
							}
						}
					}
					newData.add(newMap);
				}
				if(newData.size()>0){// 补全Sheet Name
					Map<String, Object> firstMap = newData.get(0);
					String actualSeq = (String) firstMap.get(SEQUENCE_FIELD);
					String headerSeq = (String) headerMap.get(SEQUENCE_FIELD);
					if(!actualSeq.equalsIgnoreCase(headerSeq)){
						newData.add(0, headerMap);
					}
				}
				newDatas.add(newData);
			}
		}
		return newDatas;
	}

	/**
	 * 获得真正的row数：<br/>
	 * <li>根据Test Case ID，整行数据确定真正的行数</li>
	 * 
	 * @param sheet
	 * @param field
	 * @return
	 */
	public int getRealRowNum(Sheet sheet, List<CellRangeAddress> mergeList) throws Exception {
		int realRow = 0;
		int i = 1;
		int merge = getMergeRow(mergeList);
		i = i + merge;//如果有合并单元格，加上
		int titleCount = 1 + merge;
		for (; i <= sheet.getLastRowNum(); i++) {
			Row currentRow = sheet.getRow(i);
			if(currentRow == null || "".equals(currentRow.toString())){
				break;
			}
			realRow = i + 1;
		}
		return (realRow - titleCount);
	}
	
	/**
	 * Description 判断列头是否有合并单元格
	 * @param sheet
	 */
	public Integer getMergeRow(List<CellRangeAddress> mergeList ) {
		int merge = 0;
		if(mergeList!=null && !mergeList.isEmpty()) {
			for(CellRangeAddress range : mergeList) {
				int firstRow = range.getFirstRow();
				int lastRow = range.getLastRow();
				int firstCell = range.getFirstColumn();
				cellRangeMap.put(firstRow + SPERATOR + firstCell, range);
				if(firstRow == 0 && lastRow>0) {
					if(merge< ( lastRow - firstRow) ) {
						merge = lastRow - firstRow;
					}
				}
			}
		}
		return merge;
	}

	/**
	 * 获得真正的column数
	 * 
	 * @param sheet
	 * @return
	 */
	public int getRealColNum(Sheet sheet) {
		int num = 0;
		Row headRow = sheet.getRow(0);
		Row secondRow = sheet.getRow(1);
		num = headRow.getLastCellNum();
		if(num < secondRow.getLastCellNum()) {
			num = secondRow.getLastCellNum();
		}
		return num;
	}
	
	
	/**
	 * Description 校验下拉框输入
	 * @return
	 * @throws APIException 
	 */
	public String checkPickVal(String header, String field, String value, MKSCommand cmd) throws APIException{
		if(value == null || "".equals(value)){
			return null;
		}
		List<String> valList = PICK_FIELD_RECORD.get(field);
		if(valList == null){
			valList = cmd.getAllPickValues(field);
		}
		if(valList == null ){
			return "Column [" + (header!=null?header:field )  + "] has no valid option value!";
		}else if( !valList.contains(value)){
			return "Value [" +value+ "] is invalid for Column [" + (header!=null?header:field )  + "], valid values is " + Arrays.toString(valList.toArray()) + "!";
		}
		return null;
	}
	
	/**
	 * Description 校验关联字段输入
	 * @return
	 */
	public String checkRelationshipVal(){
		
		return "";		
	}
	
	/**
	 * Description 校验用户输入
	 * @return
	 */
	public String checkUserVal(String value , String field){
		int leftIndex = -1;
		int rightIndex = -1;
		boolean endFormat = false;
		if(value.indexOf("(") > -1 ) {
			leftIndex = value.indexOf("(");
		}else if(value.indexOf("（") > -1){
			leftIndex = value.indexOf("（");
		}
		if(value.indexOf(")") > -1 ){
			rightIndex = value.indexOf(")");
			endFormat = value.endsWith(")");
		}else if(value.indexOf("）") > -1){
			rightIndex = value.indexOf("）");
			endFormat = value.endsWith("）");
		}
		String formatValue = null;
		if( leftIndex > 0 && rightIndex > 0 && endFormat) {
			formatValue = value.substring(leftIndex+1 , rightIndex );
		}else{
			formatValue = value;
		}
		if( USER_FULLNAME_RECORD.contains( formatValue.toLowerCase() ) ) {
			IS_USER = true; // 若用户存在修改标识 ， 往下执行好判断
			return "";
		}
		return "Column ["+ field +"] input value ["+ value +"] is not exist";
	}
	
	/**
	 * Description 校验relationship 输入的ID 是否带[]，是的话去掉
	 * @return
	 */
	public String checkRelationshipVal(String value){
		if(value.startsWith("[") && value.endsWith("]")){
			RELATIONSHIP_MISTAKEN = true;
		}
		return "";
	}
	
	/**
	 * Description 校验组输入
	 * @return
	 */
	public String checkGroupVal(){
		
		return "";
	}
	
	/**
	 * Description 校验组输入
	 * @return
	 */
	public String checkBooleanVal(){
		
		return "";
	}
	
	/**
	 * Description 校验输入值是否合法
	 * @return
	 * @throws APIException 
	 */
	public String checkFieldValue(String header, String field, String value, MKSCommand cmd) throws APIException{
		String fieldType = FIELD_TYPE_RECORD.get(field);
		
		if("pick".equalsIgnoreCase(fieldType)){
			return checkPickVal(header, field, value, cmd);
		}
		if("Category".equalsIgnoreCase(field)){
			return checkCategory(value);
		}
		if("Date".equalsIgnoreCase(fieldType)){
			return checkDate(value);
		}
		if("User".equalsIgnoreCase(fieldType)){
			return checkUserVal(value , field);
		}
		if("relationship".equalsIgnoreCase(fieldType)){ 
			return checkRelationshipVal(value); // 检查关联的ID是不是带 []  
		}
		return null;
	}
	
	/**
	 * Description 校验Category
	 * @return
	 */
	public String checkCategory(String value){
		if(!CURRENT_CATEGORIES.contains(value)){
			return "[" +value+ "] is invalid for Category, valid values is " + Arrays.toString(CURRENT_CATEGORIES.toArray()) + "!";
		}
		return null;
	}
	
	/**
	 * Description 校验时间格式
	 * @return
	 */
	public String checkDate(String value){
		value=value.trim();
		SimpleDateFormat sdf2 = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
		SimpleDateFormat sdf3 = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss");
		Date date = null;
		try {
			date = sdf2.parse(value);
			if(date == null)
				date = sdf3.parse(value);
		} catch (ParseException e) {
			e.printStackTrace();
		}
		if(date == null) {
			return "[" +value+ "] input error, The date and date you entered is incorrectly formatted."
								+"The Correct Format : [yyyy-MM-dd HH:mm:ss] [yyyy/MM/dd HH:mm:ss] " ;
		}
		return null;

	}
	
	/**
	 * Description 处理数据，并校验
	 * @param data
	 * @param importType
	 * @param cmd
	 * @return
	 * @throws Exception 
	 */
	@SuppressWarnings({ "unchecked", "static-access" })
	public List<List<Map<String, Object>>> checkExcelData(List<List<Map<String, Object>>> datas,Map<String,String> errorRecord, MKSCommand cmd) throws Exception {
		List<List<Map<String, Object>>> resultDatas = new ArrayList<>();
		List<Map<String, Object>> resultData = null;
		TestCaseImport.logger.info("Begin Deal Excel Sheet ,Total Sheet is :" + datas.size() );
		if(datas != null && !datas.isEmpty()){
			StringBuffer allMessage = new StringBuffer();
			List<String> sesssionIds = new ArrayList<>();
			for(int sheet=0; sheet<datas.size(); sheet++){
				resultData = new ArrayList<>();
				List<Map<String, Object>> data = datas.get(sheet);
				TestCaseImport.logger.info("Begin Deal Sheet " + (sheet + 1) + " ,Total Data in Sheet is :" + data.size() );
				if(FIELD_TYPE_RECORD == null || FIELD_TYPE_RECORD.isEmpty()){
					/** 查询Field ，为Field校验做准备*/
					List<String> importFields = new ArrayList<String>();
					for (String header : caseFields) {
						Map<String, String> fieldConfig = headerConfig.get(header);
						if(fieldConfig != null) {
							String field = fieldConfig.get("field");
							if(!"-".equals(field)){
								importFields.add(field);
							}
						}
					}
					FIELD_TYPE_RECORD.putAll(cmd.getAllFieldType(importFields,PICK_FIELD_RECORD));
				}
				if(CURRENT_CATEGORIES.isEmpty()){
					parseCurrentCategories(IMPORT_DOC_TYPE);
				}
				this.USER_FULLNAME_RECORD.addAll( cmd.getAllUserIdAndName() ); // 查询出所有的user的name 和  Id 然后存放在 USER_FULLNAME_RECORD 
				Map<String, Object> newMap = null;
				for(int i = 0; i < data.size(); i++ ) {
					boolean hasError = false;//校验出错误
					StringBuffer errorMessage = new StringBuffer();
					Map<String,Object> rowMap = data.get(i);
					String caseID = (String)rowMap.get(TEST_CASE_ID);
					newMap = new HashMap<String,Object>();
					if(caseID != null && !"".equals(caseID)) {
						newMap.put("ID", caseID);
					}
					for (String header : caseFields) {
						Map<String, String> fieldConfig = headerConfig.get(header);
						if(fieldConfig != null) {
							String field = fieldConfig.get("field");
							String value = (String)rowMap.get(field);
							if (!"-".equals(field) && value != null && !"".equals(value)) {
								String message = checkFieldValue(header, field, value, cmd);//校验Test Case字段值
								if(message == null || "".equals(message)){
									// 在此已经判断用户是否存在  ， 若存在 IS_USER 标识为 ture , 若不存在为 false
									if(IS_USER) {
										// list.get(p).toString()			
										// 判断导入的user类型的数据格式是不是 : 用户(ID) 是的话截取 ()内ID 。
										int leftIndex = -1;
										int rightIndex = -1;
										boolean endFormat = false;
										if(value.indexOf("(") > -1 ) {
											leftIndex = value.indexOf("(");
										}else if(value.indexOf("（") > -1){
											leftIndex = value.indexOf("（");
										}
										if(value.indexOf(")") > -1 ){
											rightIndex = value.indexOf(")");
											endFormat = value.endsWith(")");
										}else if(value.indexOf("）") > -1){
											rightIndex = value.indexOf("）");
											endFormat = value.endsWith("）");
										}
										if( leftIndex > 0 &&  rightIndex > 0 && endFormat ){
											String userId = value.substring(leftIndex+1 , rightIndex ); 
											if( userId.matches("d{0,9}") || userId.matches("d{0,9}") 
													|| userId.matches("d{0,9}") || userId.matches("d{0,9}") ){
												// 判断里面ID格式是不是 GW + 数字  是的话在之前查询的数据获取值
												newMap.put(field, userId);
											}
											
										}else if( value.matches("d{0,9}") || value.matches("d{0,9}")
												|| value.matches("d{0,9}") || value.matches("d{0,9}")){ // 判断如果不是用户(ID)的格式 , 在判断是不是直接填写的ID GW+数字 格式。
											
											newMap.put(field, value);
											
										}else {
											errorMessage.append(" Field ["+ field +"]  data format should be \"name(Login ID)\" or \"Login ID\" \n");
											hasError = true;
										}
										IS_USER = false;
									}else if( RELATIONSHIP_MISTAKEN ){ //如果是Relationship类型的字段，并且数字前面带[] ，就将中括号去掉
										value = value.substring(1,value.length()-1);// 
										newMap.put(field, value);
										RELATIONSHIP_MISTAKEN = false;
									}else {
										newMap.put(field, value);
									}
									
								}else{
									errorMessage.append("line "+ (i+3) +": ").append(message).append("\n");
									hasError = true;
								}
							}
						}
					}
					if(hasError){
						allMessage.append(errorMessage);
						continue;
					}
					if(rowMap.containsKey(TEST_STEP)) {//Test Case包含有 Test Step信息
						Object steps = rowMap.get(TEST_STEP);
						if(steps instanceof List){
							List<Map<String, String>> currentSteps = (List<Map<String, String>> )steps;
							if(!currentSteps.isEmpty()) {//Test Case包含有 Test Step信息
								for(Map<String, String> stepMap : currentSteps){
									for(String header : stepFields){
										Map<String, String> fieldConfig = headerConfig.get(header);
										String fieldVal = fieldConfig.get("field");
										if(fieldVal != null){
											String value = stepMap.get(header);
											stepMap.remove(header);
											stepMap.put(fieldVal, value);
										}
									}
								}
								newMap.put(TEST_STEP, currentSteps);
							}
						}
					}
					if(rowMap.containsKey(TEST_RESULT)) {//Test Case包含有 Test Result信息
						Object steps = rowMap.get(TEST_RESULT);
						if(steps instanceof List){
							List<Map<String, String>> currentResults = (List<Map<String, String>> )steps;
							if(!currentResults.isEmpty()) {//Test Case包含有 Test Result信息
								for(Map<String, String> map : currentResults) {//循环校验Test Result信息
									String sessionID = map.get(SESSION_ID);
									if(sessionID == null || "".equals(sessionID))
										allMessage.append("line "+ (i+3) +": Session ID is Empty for Import Test Result! \n");
									else
										sesssionIds.add(sessionID);
									String verdict = map.get(VERDICT);
									if(verdict == null || "".equals(verdict))
										allMessage.append("line "+ (i+3) +": verdict is Empty for Import Test Result! \n");
								}
								newMap.put(TEST_RESULT, currentResults);
							}
						}
					}
					resultData.add(newMap);
				}
				TestCaseImport.logger.info("End Deal Excel Sheet ,Total Sheet is :" + datas.size() );
				resultDatas.add(resultData);
			}
			allMessage.append(cmd.checkIssueType(sesssionIds, TEST_SESSION, SESSIONN_STATE));
			errorRecord.put("error", allMessage.toString());
		}
		return resultDatas;
	}

	/**
	 * Description 开始导入数据
	 * @param data
	 * @param cmd
	 * @param importType
	 * @param shortTitle
	 * @param project
	 * @param testSuiteID
	 * @throws Exception
	 */
	@SuppressWarnings("unchecked")
	public void startImport(List<List<Map<String, Object>>> datas, MKSCommand cmd, String importType,String shortTitle, String project, String testSuiteID) throws Exception {
		// 删除Token
		// TestCaseImport.TOKEN = null;
		// 下面List用于收集操作信息，用于统计
		List<String> caseUpdate = new ArrayList<String>(), caseCreate = new ArrayList<String>(),
				stepUpdate = new ArrayList<String>(), stepCreate = new ArrayList<String>();
		List<String> caseUpdateF = new ArrayList<String>(), caseCreateF = new ArrayList<String>(),
				stepUpdateF = new ArrayList<String>(), stepCreateF = new ArrayList<String>();

//		int totalSheetNum = data.size();
		boolean hasStep = false;
		// 遍历信息
		
		boolean createTest = false;
		if (datas.isEmpty()) {
			return;
		}
		if(testSuiteID == null || "".equals(testSuiteID)){
			Map<String,String> docInfo = new HashMap<String,String>();
			docInfo.put("Document Short Title", shortTitle);
			docInfo.put("Project", project);
			docInfo.put("State", "Open"); 
			if(IMPORT_DOC_TYPE.endsWith("Document"))
				docInfo.put("Shared Category", "Document");
			else if("Test Suite".equals(IMPORT_DOC_TYPE))
				docInfo.put("Shared Category", "Suite");
			testSuiteID = cmd.createDocument(IMPORT_DOC_TYPE, docInfo);
			createTest = true;
		}
		if (!createTest) {
			project = cmd.getItemByIds(Arrays.asList(testSuiteID), Arrays.asList("Project")).get(0).get("Project");
		}
		String parentId = testSuiteID;//涉及
		List<Map<String, Object>> data = null;
		Map<String,String> structureRecord = null;
		for(int sheet = 0; sheet<datas.size(); sheet++){
			data = datas.get(sheet);
			structureRecord = new HashMap<String,String>();
			TestCaseImport.logger.info("Start to deal sheet : " + (sheet + 1));
			int totalCaseNum = data.size();
			TestCaseImport.logger.info("Start Import Sheet " + (sheet + 1) + " Data , all Data size is :" + totalCaseNum );
			for (int index=0; index<totalCaseNum ; index++) {
				Map<String, Object> testCaseData = data.get(index);
				logger.info("Now Deal row " + index + " data");
				int caseNum = index + 1;
				String caseId = null;
				if (testCaseData.containsKey("ID")) {
					caseId = testCaseData.get("ID").toString();
				}
				if(caseId == null || "".equals(caseId)) {
					TestCaseImport.showLogger(" \tStart to Create "+ importType);
				}else {
					TestCaseImport.showLogger(" \tStart to deal "+ importType +"  : " + caseId);
				}
				Map<String, String> newTestCaseData = new HashMap<>();
				List<String> newRelatedStepIds = new ArrayList<>();
				// 1. 先处理Test
				// Step信息(更新创建或删除)，遍历得到OPERATING_ACTION和EXPECTED_RESULTS信息塞入newTestCaseData中
				if (testCaseData.containsKey(TEST_STEP)) {
					this.getTestStep(newTestCaseData, newRelatedStepIds, testCaseData, project, cmd, stepCreate,
							stepCreateF, stepUpdate, stepUpdateF);
					hasStep = true;
				}
				// 把Test Result信息获取出来
				List<Map<String,String>> resultList = null;
				if(testCaseData.get(TEST_RESULT) != null ){
					resultList = (List<Map<String,String>>)testCaseData.get(TEST_RESULT);
					testCaseData.remove(TEST_RESULT);
				}
				
 				// 2. 再处理Test Case的信息(更新或创建，不包括创建)
				String beforeId = "last";//涉及结构
				parentId = testSuiteID;
				String sequenceStr = (String)testCaseData.get(SEQUENCE_FIELD);
				String parentSeqence = null;
				if(sequenceStr != null && !"".equals(sequenceStr) ){
					if(sequenceStr.contains(SPERATOR)){
						parentSeqence = sequenceStr.substring(0, sequenceStr.lastIndexOf(SPERATOR));
						parentId = structureRecord.get(parentSeqence);
					}
					if(parentSeqence!=null)
						beforeId = structureRecord.get(parentSeqence + CHILD);
					if(beforeId == null)
						beforeId = "last";
				}
				caseId = this.getTestCase(parentId, newTestCaseData, testCaseData, project, cmd, caseId, beforeId,
						caseCreate, caseCreateF, caseUpdate, caseUpdateF, importType);
				testCaseData.put("ID", caseId);
				// 3. 关联Test Case与Test Step
				if(testCaseData.containsKey(TEST_STEP) && newRelatedStepIds.size() >0){
						this.relatedCaseAndStep(caseId, newRelatedStepIds, cmd);
				}
				// 4. 记录beforeID及结构
				structureRecord.put(sequenceStr, caseId);
				if(parentSeqence != null )
					structureRecord.put(parentSeqence + CHILD, caseId);
				// 5. 导入测试结果
				dealTestResults(resultList, cmd, caseId);

				TestCaseImport.showProgress(1, 1, caseNum, totalCaseNum);
			}
			TestCaseImport.logger.info("Success to Import sheet : " + (sheet + 1) + ". " );
		}
		
		TestCaseImport.showLogger("End to deal "+ importType +" : " + testSuiteID);
		TestCaseImport.showLogger("==============================================");
		TestCaseImport.showLogger("Create "+ CONTENT_TYPE +": success (" + caseCreate.size() + "," + caseCreate + "), failed ("
				+ caseCreateF.size() + ")");
		TestCaseImport.showLogger("Update "+ CONTENT_TYPE +": success (" + caseUpdate.size() + "," + caseUpdate + "), failed ("
				+ caseUpdateF.size() + "," + caseUpdateF + ")");
		if(hasStep) {
			TestCaseImport.showLogger("Create Test Step: success (" + stepCreate.size() + "," + stepCreate + "), failed ("
					+ stepCreateF.size() + ")");
			TestCaseImport.showLogger("Update Test Step: success (" + stepUpdate.size() + "," + stepUpdate + "), failed ("
					+ stepUpdateF.size() + "," + stepUpdateF + ")");
		}
	}

	/**
	 * 将Test Case与Test Step的关联关系进行更新
	 * 
	 * @param caseId
	 * @param newRelatedStepIds
	 * @param cmd
	 * @throws APIException
	 */
	public void relatedCaseAndStep(String caseId, List<String> newRelatedStepIds, MKSCommand cmd) throws APIException {
		if (caseId != null && caseId.length() > 0) {
			StringBuffer sb = new StringBuffer();
			for (String step : newRelatedStepIds) {
				sb.append(sb.toString().length() > 0 ? "," + step : step);
			}
			Map<String, String> map = new HashMap<>();
			map.put("Test Steps", sb.toString());
			cmd.editissue(caseId, map);
		}
	}

	/**
	 * 创建或更新Test Case
	 * 
	 * @param documentId
	 *            Suite ID
	 * @param newTestCaseData
	 *            新的Case信息集合
	 * @param caseMap
	 *            原有的Case信息集合
	 * @param project
	 *            Suite的Project
	 * @param cmd
	 * @param caseId
	 * @param beforeId
	 * @param caseCreate
	 * @param caseCreateF
	 * @param caseUpdate
	 * @param caseUpdateF
	 * @throws Exception 
	 */
	public String getTestCase(String parentId, Map<String, String> newTestCaseData, Map<String, Object> caseMap,
			String project, MKSCommand cmd, String caseId, String beforeId, List<String> caseCreate,
			List<String> caseCreateF, List<String> caseUpdate, List<String> caseUpdateF, String importType) throws Exception {
		
		logger.info("Data Of " + CONTENT_TYPE + " ID [" + caseId + "]");
		// 需修改
		for (Map.Entry<String, Object> entrty : caseMap.entrySet()) {
			String field = entrty.getKey();
			Object value = entrty.getValue();
			if (value != null && value.toString().length() > 0) {
				newTestCaseData.put(field, value.toString());
			}
		}
		String containedBy = newTestCaseData.get("Contained By");
		newTestCaseData.remove("ID");
		newTestCaseData.remove("Document ID");
		newTestCaseData.remove("Test Step");
		newTestCaseData.remove("Contained By");
		if (caseId == null || caseId.length() == 0) {
			// 创建Test Case
			try {
				if(containedBy!=null && !"".equals(containedBy) && containedBy.matches("[0-9]*")){
					parentId = containedBy;
				}
				newTestCaseData.put("Project", project);
				newTestCaseData.put("State", INIT_CONTENT_STATE);
				caseId = cmd.createContent(parentId, newTestCaseData, CONTENT_TYPE, beforeId);
				caseCreate.add(caseId);
				TestCaseImport.showLogger(" \tSuccess to create "+ CONTENT_TYPE +" : " + caseId);
			} catch (APIException e) {
				caseCreateF.add(caseId);
				TestCaseImport.showLogger(" \tFailed to create "+ CONTENT_TYPE +" : " + caseId);
				logger.error("Failed to create test case : " + ExceptionUtil.catchException(e));
			}
		} else {
			// 更新Test Case
			// 遍历出所有 overRide为 true 的字段，
			Map<String, Map<String, String>> fieldMaps= headerConfig;
			Collection<Map<String, String>> fieldMapValues = fieldMaps.values();
			List <String> fields = new ArrayList <String>();
			for( Map<String, String> values : fieldMapValues) {
				if( values.get("overRide").equals( "false") ) {
					fields.add( values.get("field") );
				}
			}
			// 然后调用 mks命令查询出导入的 所有 ids 的内容。判断当前为true字段是否有值 , getItemByIds(List<String> ids,List<String> field) 此方法通过Id 获取字段的值
			List <String> ids = new ArrayList <String>();
			ids.add(caseId);
			List<Map<String, String>> data = cmd.getItemByIds(ids,fields);
			Map<String,String> dataMap = data.get(0);
			for (String field : fields) {
				String fieldValue = dataMap.get(field);
				// 有  ： 不更新       没有 ： 更新      
				if(!"".equals(fieldValue) && null != fieldValue ){
					newTestCaseData.remove(field);
				} 
			}
			// 判断当前条目中是否 含有 Text 字段，如果有，检查此字段是否可以编辑更新（含有Text字段的条目，是否可以更新，在XML里有属性OnlyCreate 规定  。false为可编辑，true为不可编辑）
			checkOnlyCreate(newTestCaseData, importType);
			try {
				cmd.editissue(caseId, newTestCaseData);
				caseUpdate.add(caseId);
				// 1.更新顺序
				if (beforeId != null && !"".equals(beforeId)) {
					cmd.moveContent(parentId, beforeId, caseId);
				}
				TestCaseImport.showLogger(" \tSuccess to update Test Case : " + caseId);
			} catch (APIException e) {
				caseUpdateF.add(caseId);
				TestCaseImport.showLogger(" \tFailed to update Test Case : " + caseId);
				logger.error("Failed to edit test case : " + ExceptionUtil.catchException(e));
			}
		}
		return caseId;
	}
	/**
	 * 	检测当前要更新Case里面有没有 Text 字段 ， 并且判断该字段是否可以编辑        xml 中有 onlyCreate 属性规定是否可以更新
	 * @param newTestCaseData
	 * @param importType
	 */
	private void checkOnlyCreate(Map<String, String> newTestCaseData, String importType) {
		Collection<Map<String, String>> values= headerConfig.values();
		for (Map<String, String> map : values) {
			if(map.get("onlyCreate") != null) {
				boolean onlyCreate = Boolean.valueOf(map.get("onlyCreate"));
				if(onlyCreate) {
					String field = map.get("field");
					newTestCaseData.remove(field);
				}
			}
		}
	}
	/**
	 * 先处理Test Step信息(更新创建或删除)，
	 * 遍历得到OPERATING_ACTION和EXPECTED_RESULTS信息塞入newTestCaseData中, 并将创建和更新的Step
	 * ID塞于newRelatedStepIds中
	 * 
	 * @param newTestCaseData
	 *            新的Case信息集合
	 * @param newRelatedStepIds
	 *            创建Test Step的集合
	 * @param caseMap
	 *            原有Case信息集合
	 * @param project
	 *            Suite的Project信息
	 * @param cmd
	 * @param stepCreate
	 * @param stepCreateF
	 * @param stepUpdate
	 * @param stepUpdateF
	 */
	public void getTestStep(Map<String, String> newTestCaseData, List<String> newRelatedStepIds,
			Map<String, Object> caseMap, String project, MKSCommand cmd, List<String> stepCreate,
			List<String> stepCreateF, List<String> stepUpdate, List<String> stepUpdateF) {
		@SuppressWarnings("unchecked")
		List<Map<String, String>> testStepData = (List<Map<String, String>>) caseMap.get(TEST_STEP);
		int i = 1;
		if (testStepData != null && testStepData.size() > 0) {
			TestCaseImport.showLogger(" \t\tHas Test Step size  : " + testStepData.size());
			for (Map<String, String> stepMap : testStepData) {
//				if (i == 1) {// 将第一步的Test Step的Precondition加入到Test Case中
//					newTestCaseData.put(INITIAL_STATE_PRECODITION,
//							stepMap.get(PRECODITION) == null ? "" : stepMap.get(PRECODITION));
//				}
				String stepId = stepMap.get("ID");
				stepMap.remove("ID");
				// 处理Step Order
//				stepMap.put(STEP_ORDER, i + "");
//				if (stepMap.get("Test Step") == null || stepMap.get("Test Step").trim().length() == 0) {
//					stepMap.put("Test Step", i + "");
//				}
				if (stepId == null || stepId.length() == 0) {
					// 创建Test Step，并关联Test Case
					try {
						stepMap.put("Project", project);
						stepMap.put("State", INIT_STEP_STATE);
						stepId = cmd.createIssue(TEST_STEP, stepMap, null);
						stepCreate.add(stepId);
						TestCaseImport.showLogger(" \t\tSuccess to create Test Step " + i + ", " + stepId);
					} catch (APIException e) {
						stepCreateF.add(stepId);
						TestCaseImport.showLogger(" \t\tFailed to create Test Step");
						logger.error("Failed to create test step : " + ExceptionUtil.catchException(e));
					}
				} else {
					try {
						cmd.editissue(stepId, stepMap);
						stepUpdate.add(stepId);
						TestCaseImport.showLogger(" \t\tSuccess to update Test Step " + i + ", " + stepId);
					} catch (APIException e) {
						stepUpdateF.add(stepId);
						TestCaseImport.showLogger(" \t\tFailed to update Test Step " + i + ", " + stepId);
						logger.error("Failed to edit test step : " + ExceptionUtil.catchException(e));
					}
				}
				newRelatedStepIds.add(stepId);
				i++;
			}
		}
	}
	
	/**
	 * 处理导入结果
	 * @param caseMap
	 * @param cmd
	 */
	public void dealTestResults(List<Map<String, String>> resultDatas, MKSCommand cmd, String caseID){
		if (resultDatas != null && !resultDatas.isEmpty()) {
			for (Map<String, String> result : resultDatas) {
				String sessionId = result.get(SESSION_ID);
				String verdict = result.get(VERDICT);
				String annotation = result.get("Annotation");
				cmd.createResult(sessionId, verdict, annotation, caseID);
			}
		}
	}
}
