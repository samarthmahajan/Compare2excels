package com.excel.utility.excelcomparison.dto;

import java.util.List;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

@NoArgsConstructor
@Data
@AllArgsConstructor
@Builder
public class Result {

	String sheetName;
	List<String> errorsFould;
	boolean equal;
}
