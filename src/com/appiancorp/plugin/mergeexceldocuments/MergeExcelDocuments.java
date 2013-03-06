package com.appiancorp.plugin.mergeexceldocuments;

import org.apache.log4j.Logger;


import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.StringWriter;
import java.io.IOException;
import java.sql.Timestamp;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.HashMap;
import java.util.Locale;
import java.util.Map;
import java.util.ResourceBundle;
import java.util.TreeSet;

import org.apache.commons.io.IOUtils;
import org.apache.commons.lang.StringUtils;
import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.util.CellRangeAddress;


import com.appiancorp.common.SystemMessageResolver;
import com.appiancorp.suiteapi.common.Name;
import com.appiancorp.suiteapi.common.ObjectTypeMapping;
import com.appiancorp.suiteapi.content.Content;
import com.appiancorp.suiteapi.content.ContentConstants;
import com.appiancorp.suiteapi.content.ContentFilter;
import com.appiancorp.suiteapi.content.ContentService;
import com.appiancorp.suiteapi.knowledge.Document;
import com.appiancorp.suiteapi.knowledge.DocumentDataType;
import com.appiancorp.suiteapi.knowledge.FolderDataType;
import com.appiancorp.suiteapi.process.ProcessExecutionService;
import com.appiancorp.suiteapi.process.analytics2.Column;
import com.appiancorp.suiteapi.process.analytics2.Filter;
import com.appiancorp.suiteapi.process.analytics2.ProcessAnalyticsService;
import com.appiancorp.suiteapi.process.analytics2.ProcessReport;
import com.appiancorp.suiteapi.process.analytics2.ReportData;
import com.appiancorp.suiteapi.process.analytics2.ReportResultPage;
import com.appiancorp.suiteapi.process.analytics2.SimpleColumnFilter;
import com.appiancorp.suiteapi.process.exceptions.SmartServiceException;
import com.appiancorp.suiteapi.process.framework.AppianSmartService;
import com.appiancorp.suiteapi.process.framework.Input;
import com.appiancorp.suiteapi.process.framework.MessageContainer;
import com.appiancorp.suiteapi.process.framework.Required;
import com.appiancorp.suiteapi.process.framework.SmartServiceContext;
import com.appiancorp.suiteapi.process.palette.PaletteInfo;
import com.appiancorp.suiteapi.type.TypeService;
import com.appiancorp.suiteapi.type.TypedValue;

import com.appiancorp.suiteapi.process.palette.PaletteInfo; 

@PaletteInfo(paletteCategory = "Integration Services", palette = "Connectivity Services") 
public class MergeExcelDocuments extends AppianSmartService {

	private static final Logger LOG = Logger
			.getLogger(MergeExcelDocuments.class);
	private final SmartServiceContext smartServiceCtx;
	private Long[] listDocuments;
	private Long destinationFolder;
	private String combinedDocName;
	private boolean excludeHeaderRowExceptFirstDoc;
	private Long combinedExcelDocument;

  private ProcessAnalyticsService pas;
  private ContentService cs;
  private ProcessExecutionService pes;
  private TypeService ts;

  public MergeExcelDocuments(TypeService ts, ContentService cs,
      ProcessExecutionService pes, ProcessAnalyticsService pas,
      SmartServiceContext smartServiceCtx) {
    super();
    
    this.smartServiceCtx = smartServiceCtx;
    this.cs = cs;
    this.pes = pes;
    this.ts = ts;
    this.pas = pas;
  }
	  
	@Override
	public void run() throws SmartServiceException {
		// TODO Auto-generated method stub

	    Locale currentLocale = smartServiceCtx.getUserLocale() != null ? smartServiceCtx
	        .getUserLocale() : smartServiceCtx.getPrimaryLocale();

	        try {
	      
	        	XSSFWorkbook book = new XSSFWorkbook();
	            
	            ArrayList<FileInputStream> listFileInputStream = new ArrayList<FileInputStream>();
	            
	            if (listDocuments != null) {
	              for(int i=0;i<listDocuments.length;i++)
	              {
	            	  
	            	Document xls = cs.download(listDocuments[i],
	                  ContentConstants.VERSION_CURRENT, true)[0];

	            	String extension = xls.getExtension();
//	            	System.out.println("MERGE EXCEL DOC Doc Extension: " + extension);		        
	            	
	            	if(extension.equals("xlsx"))
	            	{	
	       
//	            		System.out.println("MERGE EXCEL DOC Document ID: " + listDocuments[i]);
//	            		LOG.debug("MEDOC Document ID: " + listDocuments[i]);
		              String documentPath = xls.getInternalFilename();
		              FileInputStream fis = new FileInputStream(documentPath);
		              listFileInputStream.add(fis);
	            	}
	              }

	            	if(listFileInputStream.size() > 0)
	            	{	
	  	                book = mergeExcelFiles(book, listFileInputStream);
			            
	  	                Document outputDoc = registerDocument();
			            
			            File file = new File(outputDoc.getInternalFilename());
			            FileOutputStream fos = new FileOutputStream(file);
			            
			            book.write(fos);
			            
			            fos.close();
			            
			            combinedExcelDocument = outputDoc.getId();
	            	}
	            	
	            }
            
	  // Set combinedExcelDocument to the final excel doc         
	   } catch (Exception e) {
		      LOG.error(e, e);
		      throw createException(e, "error.export.general", e.getMessage());
		    }
	        
	}
	
	 private SmartServiceException createException(Throwable t, String key,
		      Object... args) {
		    return new SmartServiceException.Builder(getClass(), t).userMessage(key,
		        args).build();
		  }	

	 public XSSFWorkbook mergeExcelFiles(XSSFWorkbook book, ArrayList<FileInputStream> inList) throws IOException {

		 	XSSFSheet finalSheet = book.createSheet();
		    for (FileInputStream fin : inList) {
		        XSSFWorkbook wb = new XSSFWorkbook(fin);
//            	System.out.println("MERGE EXCEL DOC Starting CopySheets ");
//            	LOG.debug("MEDOC Starting CopySheets ");
		        XSSFSheet sheet = wb.getSheetAt(0);
		        int lastRowNum = sheet.getLastRowNum();
//            	System.out.println("MERGE EXCEL DOC Sheet LastRowNum: " + lastRowNum);		        
		        if(lastRowNum > 0)
		        {	
		          copySheets(finalSheet, sheet, excludeHeaderRowExceptFirstDoc);
		        }
		        fin.close();
		    }
		    return book;
	 }
		    
	 public static void copySheets(XSSFSheet newSheet, XSSFSheet sheet, boolean excludeHeaderRowExceptFirst){     
		        copySheets(newSheet, sheet, excludeHeaderRowExceptFirst, false);     
		    }     
		    
	 public static void copySheets(XSSFSheet newSheet, XSSFSheet sheet, boolean excludeHeaderRowExceptFirst, boolean copyStyle){     
		    int maxColumnNum = 0;
		    int destLastRowNum = newSheet.getLastRowNum();
		    int destNextRowNum = 0;
		    boolean skipHeaderRow = false;
		    Map<Integer, XSSFCellStyle> styleMap = (copyStyle) ? new HashMap<Integer, XSSFCellStyle>() : null;		    
		    for (int i = sheet.getFirstRowNum(); i <= sheet.getLastRowNum(); i++) {     
		        XSSFRow srcRow = sheet.getRow(i);
		        destLastRowNum = newSheet.getLastRowNum();
//		        System.out.println("MERGE EXCEL DOC Dest LastRowNum: " + destLastRowNum);
//		        System.out.println("MERGE EXCEL DOC i: " + i);
		        destNextRowNum = (destLastRowNum==0 && i==0)? 0: destLastRowNum+1;  
		        skipHeaderRow = (excludeHeaderRowExceptFirst && i==0 && destNextRowNum!=0)? true:false;
//		        System.out.println("MERGE EXCEL DOC excludeHeaderRowExceptFirst: " + excludeHeaderRowExceptFirst);
//		        System.out.println("MERGE EXCEL DOC Dest destNextRowNum: " + destNextRowNum);
//		        System.out.println("MERGE EXCEL DOC skipHeaderRow: " + skipHeaderRow);
		        if (srcRow != null && !skipHeaderRow) {     
			        XSSFRow destRow = newSheet.createRow(destNextRowNum);
		            copyRow(sheet, newSheet, srcRow, destRow, styleMap);     
		            if (srcRow.getLastCellNum() > maxColumnNum) {     
		                maxColumnNum = srcRow.getLastCellNum();     
		            }     
		        }     
		    }     
		    for (int i = 0; i <= maxColumnNum; i++) {     
		        newSheet.setColumnWidth(i, sheet.getColumnWidth(i));     
		    }     
		} 	 
		    
	 public static void copyRow(XSSFSheet srcSheet, XSSFSheet destSheet, XSSFRow srcRow, XSSFRow destRow, Map<Integer, XSSFCellStyle> styleMap) {     
		    // manage a list of merged zone in order to not insert two times a merged zone  
		    destRow.setHeight(srcRow.getHeight());     
		    // reckoning delta rows  
		    int deltaRows = destRow.getRowNum()-srcRow.getRowNum();  
		    // pour chaque row  
		    for (int j = srcRow.getFirstCellNum(); j <= srcRow.getLastCellNum(); j++) {     
		        XSSFCell oldCell = srcRow.getCell(j);   // ancienne cell  
		        XSSFCell newCell = destRow.getCell(j);  // new cell   
		        if (oldCell != null) {     
		            if (newCell == null) {     
		                newCell = destRow.createCell(j);     
		            }     
		            // copy chaque cell  
		            copyCell(oldCell, newCell, styleMap);     
		            // copy les informations de fusion entre les cellules  
		            //System.out.println("row num: " + srcRow.getRowNum() + " , col: " + (short)oldCell.getColumnIndex());  
		            CellRangeAddress mergedRegion = getMergedRegion(srcSheet, srcRow.getRowNum(), (short)oldCell.getColumnIndex());     

		            if (mergedRegion != null) {   
		              //System.out.println("Selected merged region: " + mergedRegion.toString());  
		              CellRangeAddress newMergedRegion = new CellRangeAddress(mergedRegion.getFirstRow()+deltaRows, mergedRegion.getLastRow()+deltaRows, mergedRegion.getFirstColumn(),  mergedRegion.getLastColumn());  
		                //System.out.println("New merged region: " + newMergedRegion.toString());  
		            }     
		        }     
		    }                
		} 
	 

	 /** 
	  * @param oldCell 
	  * @param newCell 
	  * @param styleMap 
	  */  
	 public static void copyCell(XSSFCell oldCell, XSSFCell newCell, Map<Integer, XSSFCellStyle> styleMap) {     
	     if(styleMap != null) {     
	         if(oldCell.getSheet().getWorkbook() == newCell.getSheet().getWorkbook()){     
	             newCell.setCellStyle(oldCell.getCellStyle());     
	         } else{     
	             int stHashCode = oldCell.getCellStyle().hashCode();     
	             XSSFCellStyle newCellStyle = styleMap.get(stHashCode);     
	             if(newCellStyle == null){     
	                 newCellStyle = newCell.getSheet().getWorkbook().createCellStyle();     
	                 newCellStyle.cloneStyleFrom(oldCell.getCellStyle());     
	                 styleMap.put(stHashCode, newCellStyle);     
	             }     
	             newCell.setCellStyle(newCellStyle);     
	         }     
	     }     
	     switch(oldCell.getCellType()) {     
	         case XSSFCell.CELL_TYPE_STRING:     
	             newCell.setCellValue(oldCell.getStringCellValue());     
	             break;     
	       case XSSFCell.CELL_TYPE_NUMERIC:     
	             newCell.setCellValue(oldCell.getNumericCellValue());     
	             break;     
	         case XSSFCell.CELL_TYPE_BLANK:     
	             newCell.setCellType(XSSFCell.CELL_TYPE_BLANK);     
	             break;     
	         case XSSFCell.CELL_TYPE_BOOLEAN:     
	             newCell.setCellValue(oldCell.getBooleanCellValue());     
	             break;     
	         case XSSFCell.CELL_TYPE_ERROR:     
	             newCell.setCellErrorValue(oldCell.getErrorCellValue());     
	             break;     
	         case XSSFCell.CELL_TYPE_FORMULA:     
	             newCell.setCellFormula(oldCell.getCellFormula());     
	             break;     
	         default:     
	             break;     
	     }     

	 }     
	 
	 public static CellRangeAddress getMergedRegion(XSSFSheet sheet, int rowNum, short cellNum) {     
		    for (int i = 0; i < sheet.getNumMergedRegions(); i++) {   
		        CellRangeAddress merged = sheet.getMergedRegion(i);     
		        if (merged.isInRange(rowNum, cellNum)) {     
		            return merged;     
		        }     
		    }     
		    return null;     
		}     
	 
	 
	 private Document registerDocument() throws Exception {

		    String name = combinedDocName;
		    String extension = "xlsx";

		    ContentFilter cf = new ContentFilter(ContentConstants.TYPE_DOCUMENT);
		    cf.setName(name);
		    cf.setExtension(new String[] { extension });

		    Long docId = null;
	        Document d = new Document();
	        d.setName(name);
	        d.setExtension(extension);
	        d.setSize(1);
	        d.setParent(destinationFolder);
	        d.setState(ContentConstants.STATE_ACTIVE_PUBLISHED);
	        docId = cs.create(d, ContentConstants.UNIQUE_NONE);

		    if (docId != null) {
		      Document outputDoc = cs.download(docId, ContentConstants.VERSION_CURRENT, false)[0];
		      return outputDoc;
		    }

		    return null;
		  }
	 
	 
	public void onSave(MessageContainer messages) {
	}

	public void validate(MessageContainer messages) {
	}

	@Input(required = Required.ALWAYS)
	@Name("ListDocuments")
	@DocumentDataType
	public void setListDocuments(Long[] val) {
		this.listDocuments = val;
	}

	@Input(required = Required.ALWAYS)
	@Name("DestinationFolder")
	@FolderDataType
	public void setDestinationFolder(Long val) {
		this.destinationFolder = val;
	}

	@Input(required = Required.ALWAYS)
	@Name("CombinedDocName")
	public void setCombinedDocName(String val) {
		this.combinedDocName = val;
	}

	@Input(required = Required.ALWAYS)
	@Name("ExcludeHeaderRowExceptFirstDoc")
	public void setExcludeHeaderRowExceptFirstDoc(boolean val) {
		this.excludeHeaderRowExceptFirstDoc = val;
	}

	@Name("CombinedExcelDocument")
	@DocumentDataType
	public Long getCombinedExcelDocument() {
		return combinedExcelDocument;
	}

}
