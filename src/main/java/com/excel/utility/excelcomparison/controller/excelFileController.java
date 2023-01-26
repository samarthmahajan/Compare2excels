package com.excel.utility.excelcomparison.controller;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import com.excel.utility.excelcomparison.Service.ExcelAnalyserService;
import com.excel.utility.excelcomparison.dto.FinalResult;

import lombok.extern.slf4j.Slf4j;

@RestController
@RequestMapping("/compareExcel")
@Slf4j
public class excelFileController {
	
	@Value("${source.excel.file}")
	private String FILE_NAME ;
	
	@Autowired
	public ExcelAnalyserService excelAnalyserService;
	
	@PostMapping(consumes = { MediaType.MULTIPART_FORM_DATA_VALUE })
	 public ResponseEntity getCompareExcelsAndFindTheResult(@RequestParam(name = "fileToComapre") MultipartFile fileToCompare) throws Exception {
		log.info("Starting excelComparisation for File : "+ fileToCompare.getName());
		FinalResult res = excelAnalyserService.analyseExcel(fileToCompare);
		log.info("Comparisation Completed ");
		return ResponseEntity.ok(res); 
	}
	
}
