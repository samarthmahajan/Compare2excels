package com.excel.utility.excelcomparison.Service;

import java.io.FileNotFoundException;
import java.io.IOException;

import org.springframework.web.multipart.MultipartFile;

import com.excel.utility.excelcomparison.dto.FinalResult;
public interface ExcelAnalyserService {
	FinalResult analyseExcel(MultipartFile fileToCompare) throws FileNotFoundException, IOException;
}
