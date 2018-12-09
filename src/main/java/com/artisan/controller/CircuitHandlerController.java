package com.artisan.controller;

import javax.servlet.http.HttpServletResponse;
import javax.validation.Valid;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Qualifier;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import com.artisan.service.CircuitServiceInterface;
import com.artisan.util.Util;

import io.swagger.annotations.Api;
import io.swagger.annotations.ApiOperation;
import io.swagger.annotations.ApiParam;
import io.swagger.annotations.ApiResponse;
import io.swagger.annotations.ApiResponses;

@RestController
@RequestMapping("/v1")
@Api(value = "电路调度操作", tags = { "电路调度操作" })
public class CircuitHandlerController<T> {

	@Autowired
	@Qualifier("circuitService")
	private CircuitServiceInterface circuitService;

	@ApiOperation(tags = "导出工单和电路信息", value = "导出工单和电路信息", notes = "导出工单和电路信息", response = Void.class)
	@ApiResponses({ @ApiResponse(code = 0, message = "导出成功", response = Void.class),
			@ApiResponse(code = 201, message = "导出报错", response = Void.class) })
	@GetMapping("/exportFormAndCircuitInfo/{flowId}")
	public void exportFormAndCircuitInfo(
			@Valid @ApiParam(required = true, value = "http response") 
			HttpServletResponse response, @PathVariable("flowId") String flowId) {
			try {
				XSSFWorkbook workBook = this.circuitService.exportFormAndCircuitInfo(flowId);

				response.setContentType("application/vnd.ms-excel");
				response.setHeader("Content-Disposition",
						"attachment;filename=" + Util.toDownloadString("工单和电路信息"+flowId) + ".xlsx");
				
				workBook.write(response.getOutputStream());
			}catch (Exception e) {
				e.printStackTrace();
			}
		}	
}
