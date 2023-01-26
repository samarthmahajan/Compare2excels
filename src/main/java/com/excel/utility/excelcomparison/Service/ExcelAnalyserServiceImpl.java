package com.excel.utility.excelcomparison.Service;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.net.URL;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import com.excel.utility.excelcomparison.dto.FinalResult;
import com.excel.utility.excelcomparison.dto.Result;

import lombok.extern.slf4j.Slf4j;

@Slf4j
@Service
public class ExcelAnalyserServiceImpl implements ExcelAnalyserService {
	
	@Value("${source.excel.file}")
	private String FILE_NAME ;
	
	@Override
	public FinalResult analyseExcel(MultipartFile fileToCompare) throws FileNotFoundException, IOException {
	
		FileInputStream referenceFile = getReferenceFile();
		FileInputStream testExcel = getTestFile(fileToCompare);
		
		XSSFWorkbook workbook = new XSSFWorkbook(referenceFile);
        XSSFWorkbook workbookToCompare = new XSSFWorkbook(testExcel);
        
        if (workbook.getNumberOfSheets() != workbookToCompare.getNumberOfSheets())
        {
        	log.error("Files not equal.... expected no of sheets : " +  workbook.getNumberOfSheets() +" found : " + workbookToCompare.getNumberOfSheets());
        	return FinalResult.builder()
        			.conclusion("Excelsheets are not equal.... expected no of sheets : " +  workbook.getNumberOfSheets() +" found : " + workbookToCompare.getNumberOfSheets())
        			.build();
        	
        }

        FinalResult res = new FinalResult();
        boolean excelAreEqual = true;
        List<Result> sheetsRes = new ArrayList<>();
        for (int i=0 ; i<workbook.getNumberOfSheets() ; i++) {
            Result sheetRes = new Result();
        	XSSFSheet sheet1 = workbook.getSheetAt(i);
        	XSSFSheet sheet2 = workbookToCompare.getSheetAt(i);
        	sheetRes.setSheetName(sheet1.getSheetName());
        	List<String> errorsFould = new ArrayList<>();
        	boolean isEqual = compareTwoSheets(sheet1, sheet2, errorsFould);
            sheetRes.setEqual(isEqual);
            if(!isEqual) {
            	excelAreEqual = false;
            }
            sheetRes.setErrorsFould(errorsFould);
        	sheetsRes.add(sheetRes);
        }
        referenceFile.close();
        testExcel.close();
        if(excelAreEqual) {
        	res.setConclusion("Both ExcelSheets are equal");
        }else {
        	res.setConclusion("Both ExcelSheets are not equal");
        	res.setAnalysis(sheetsRes);
        }
        return res;
    
	}

	private FileInputStream getTestFile(MultipartFile fileToCompare) throws FileNotFoundException, IOException {
		File test = new File("src/main/resources/targetFile.tmp");
        try (OutputStream os = new FileOutputStream(test)) {
            os.write(fileToCompare.getBytes());
        }
       return new FileInputStream( new File("src/main/resources/targetFile.tmp"));
	}

	private FileInputStream getReferenceFile() throws FileNotFoundException {
		URL url =getClass().getClassLoader().getResource(FILE_NAME);
		if(url == null) {
			log.error("File not found ");
			throw new FileNotFoundException("Cannnot find the file");
			
		}
		 File file = new File(getClass().getClassLoader().getResource(FILE_NAME).getFile());
		  return new FileInputStream(file);
	}
	
	 public static boolean compareTwoSheets(XSSFSheet sheet1, XSSFSheet sheet2, List<String> errorsFould)
	    {
	        int firstRow1 = sheet1.getFirstRowNum();
	        int lastRow1 = sheet1.getLastRowNum();
	        boolean equalSheets = true;
	        for(int i=firstRow1; i <= lastRow1; i++)
	        {
	            System.out.print("___________________________");
	            System.out.println("\nComparing Row "+i);
	            System.out.println("___________________________");
	            XSSFRow row1 = sheet1.getRow(i);
	            XSSFRow row2 = sheet2.getRow(i);
	            if(!compareTwoRows(row1, row2, errorsFould))
	            {
	                equalSheets = false;
	                System.out.println(" Row "+i+" | Not Equal");
	            }
	            else
	            {
	                System.out.println(" Row "+i+" | Equal");
	            }
	        }
	        return equalSheets;
	        
	    }
	    
	    public static boolean compareTwoRows(XSSFRow row1, XSSFRow row2, List<String> errorsFould)
	    {
	        if((row1 == null) && (row2 == null))
	        {
	            return true;
	        }
	        else if((row1 == null) || (row2 == null))
	        {
	            return false;
	        }
	        int firstCell1 = row1.getFirstCellNum();
	        int lastCell1 = row1.getLastCellNum();
	        boolean equalRows = true;
	        List<String> error = new ArrayList<>();
	        for(int i=firstCell1; i <= lastCell1; i++)
	        {
	            XSSFCell cell1 = row1.getCell(i);
	            XSSFCell cell2 = row2.getCell(i);
	            if(!compareTwoCells(cell1, cell2))
	            {
	                equalRows = false;
	                System.err.println("Cell " +i+" | Not Equal originalValue : " + cell1.getStringCellValue() + " valueFound : " +  cell2.getStringCellValue());
	                errorsFould.add("issue found in row : "+ cell1.getRowIndex()  + " and coulumn : "+ cell1.getColumnIndex() +" | valueExpected : " + cell1.getStringCellValue() + " valueFound : " +  cell2.getStringCellValue());
	            }
	            else
	            {
	                System.out.println("Cell "+i+" | Equal");
	            }
	        }
	        return equalRows;
	    }

	    public static boolean compareTwoCells(XSSFCell cell1, XSSFCell cell2)
	    {
	        if((cell1 == null) && (cell2 == null))
	        {
	            return true;
	        }
	        else if((cell1 == null) || (cell2 == null))
	        {
	            return false;
	        }
	        
	        boolean equalCells = false;
	        int type1 = cell1.getCellType();
	        int type2 = cell2.getCellType();
	        if (type1 == type2)
	        {
	            if (cell1.getCellStyle().equals(cell2.getCellStyle()))
	            {
	                switch (cell1.getCellType())
	                {
	                    case HSSFCell.CELL_TYPE_FORMULA:
	                        if (cell1.getCellFormula().equals(cell2.getCellFormula()))
	                        {
	                            equalCells = true;
	                        }
	                        break;
	                    case HSSFCell.CELL_TYPE_NUMERIC:
	                        if (cell1.getNumericCellValue() == cell2.getNumericCellValue())
	                        {
	                            equalCells = true;
	                        }
	                        break;
	                    case HSSFCell.CELL_TYPE_STRING:
	                        if (cell1.getStringCellValue().equals(cell2.getStringCellValue()))
	                        {
	                            equalCells = true;
	                        }
	                        break;
	                    case HSSFCell.CELL_TYPE_BLANK:
	                        if (cell2.getCellType() == HSSFCell.CELL_TYPE_BLANK)
	                        {
	                            equalCells = true;
	                        }
	                        break;
	                    case HSSFCell.CELL_TYPE_BOOLEAN:
	                        if (cell1.getBooleanCellValue() == cell2.getBooleanCellValue())
	                        {
	                            equalCells = true;
	                        }
	                        break;
	                    case HSSFCell.CELL_TYPE_ERROR:
	                        if (cell1.getErrorCellValue() == cell2.getErrorCellValue())
	                        {
	                            equalCells = true;
	                        }
	                        break;
	                    default:
	                        if (cell1.getStringCellValue().equals(cell2.getStringCellValue()))
	                        {
	                            equalCells = true;
	                        }
	                        break;
	                }
	            }
	            else
	            {
	                return false;
	            }
	        }
	        else
	        {
	            return false;
	        }
	        return equalCells;
	    }


}
