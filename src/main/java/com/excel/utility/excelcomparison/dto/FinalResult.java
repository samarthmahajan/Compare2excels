package com.excel.utility.excelcomparison.dto;

import java.util.List;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

@NoArgsConstructor
@Data
@Builder
@AllArgsConstructor
public class FinalResult {
 String conclusion;
 List<Result>  analysis;
}
