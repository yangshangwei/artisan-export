package com.artisan.service.impl;

import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;
import org.springframework.util.StringUtils;

import com.artisan.dao.CircuitBusinessinfoModelMapper;
import com.artisan.dao.CircuitModelMapper;
import com.artisan.service.CircuitServiceInterface;
import com.artisan.util.ExcelUtil;
import com.artisan.util.Util;

import lombok.extern.slf4j.Slf4j;


@Service("circuitService")
@Slf4j
public class CircuitServiceImpl implements CircuitServiceInterface {
	
	@Autowired
	CircuitBusinessinfoModelMapper circuitBusinessinfoModelMapper;
	
	@Autowired
	CircuitModelMapper circuitModelMapper;
	
	@Override
	public XSSFWorkbook exportFormAndCircuitInfo(String flowId) throws Exception {
		log.info("begin to exportFormAndCircuitInfo,flowId: {} ", flowId);
		// 新建工作簿
		XSSFWorkbook workbook = new XSSFWorkbook();
		
	    // Step1:查询工单信息,生成工单信息sheet页
	    List<Map<String, Object>> circuitBSInfoList =  this.circuitBusinessinfoModelMapper.selectFormInfoByFlowId(flowId);
	    if (circuitBSInfoList!=null && !circuitBSInfoList.isEmpty()) {
	    	// 如果返回多条,脏数据,仅处理第一条
	    	Map<String, Object> circuitInfoMap = circuitBSInfoList.get(0);
	    	// 生成工单信息sheet列
	    	generateWorkOrderInfo(workbook, circuitInfoMap);
		}else {
			// 没有工单信息 生成一个空的sheet 否则下载后无法打开
			generateSheetWithoutContent(workbook);
		}
	    
	    // Step2:查询新增的电路信息,如果有数据则生成新增电路sheet页
	    List<Map<String, Object>> circuitModelAddList = this.circuitModelMapper.selectCircuitInfoByFlowIdAndAttempType(flowId,"新增");
	    circuitModelAddList.forEach(System.out::println);
	    if(circuitModelAddList!=null && !circuitModelAddList.isEmpty()){
	    	//生成新增sheet页
	    	generateCircuitSheetByAttempType(workbook, circuitModelAddList,"新增电路信息");
	    }
	    
	    // Step3:查询调整的电路信息,如果有数据则生成调整电路sheet页
	    List<Map<String, Object>> circuitModelModifyList = this.circuitModelMapper.selectCircuitInfoByFlowIdAndAttempType(flowId,"调整");
	    circuitModelAddList.forEach(System.out::println);
	    if(circuitModelModifyList!=null && !circuitModelModifyList.isEmpty()){
	    	//生成调整sheet页
	    	generateCircuitSheetByAttempType(workbook, circuitModelModifyList,"调整电路信息");
	    }
	    
	    // Step4:查询停闭的电路信息,如果有数据则生成停闭电路sheet页
	    List<Map<String, Object>> circuitModelTerminalList = this.circuitModelMapper.selectCircuitInfoByFlowIdAndAttempType(flowId,"停闭");
	    circuitModelAddList.forEach(System.out::println);
	    if(circuitModelTerminalList!=null && !circuitModelTerminalList.isEmpty()){
	    	//生成停闭sheet页
	    	generateCircuitSheetByAttempType(workbook, circuitModelTerminalList,"停闭电路信息");
	    }
		log.info("finish  exportFormAndCircuitInfo,flowId: {} ", flowId);
	    return workbook;
	}
	
	/**
	 * 当工单信息无数据时，生成默认的sheet,否则无法打卡下载的excel
	 * @param workbook
	 */
	private void generateSheetWithoutContent(XSSFWorkbook workbook) {
		workbook.createSheet("无工单内容");
	}

	/**
	 * 
	 * @param workbook
	 * @param circuitModelAddList
	 * @param sheetName
	 * @desc 根据集合和sheet名称,生成对应的sheet页
	 */
	private void generateCircuitSheetByAttempType(XSSFWorkbook workbook,
			List<Map<String, Object>> circuitModelAddList,String sheetName) {
		// 根据sheetName创建sheet页
		XSSFSheet sheet = workbook.createSheet(sheetName);
		 
		// 设置标题行的样式
		setCircuitExcelHeadRowStyle(sheet, ExcelUtil.setStyle(workbook, IndexedColors.TURQUOISE.getIndex(),false),(short) 400);
		
		// 设置数据
		for(int i=0;i<circuitModelAddList.size();i++){
			
			Map<String,Object> circuitMap = circuitModelAddList.get(i);
			XSSFRow row = sheet.createRow(i+1);
			
			row.createCell(0).setCellValue(i+1);
		    row.createCell(1).setCellValue(Util.getValueOfStr(String.valueOf(circuitMap.get("esop_order_no"))));
		    row.createCell(2).setCellValue(Util.getValueOfDate((Date)circuitMap.get("limit_time"),null));

		    row.createCell(3).setCellValue(Util.getValueOfStr(String.valueOf(circuitMap.get("attemp_type"))));
		    row.createCell(4).setCellValue(Util.getValueOfStr(String.valueOf(circuitMap.get("name"))));
		    
		    row.createCell(5).setCellValue(Util.getValueOfStr(String.valueOf(circuitMap.get("bandwidth"))));
		    row.createCell(6).setCellValue(Util.getValueOfStr(String.valueOf(circuitMap.get("rate"))));
		    row.createCell(7).setCellValue(Util.getValueOfStr(String.valueOf(circuitMap.get("qos"))));
		    row.createCell(8).setCellValue(Util.getValueOfStr(String.valueOf(circuitMap.get("a_trans_site_name"))));
		    row.createCell(9).setCellValue(Util.getValueOfStr(String.valueOf(circuitMap.get("a_trans_room_name"))));
		    row.createCell(10).setCellValue(Util.getValueOfStr(String.valueOf(circuitMap.get("start_device_name"))));
		    row.createCell(11).setCellValue(Util.getValueOfStr(String.valueOf(circuitMap.get("start_equipport_name"))));
		    row.createCell(12).setCellValue(Util.getValueOfStr(String.valueOf(circuitMap.get("a_trans_cpt_name"))));
		    
		    // TODO 14-16
		    row.createCell(13).setCellValue(Util.getValueOfStr(String.valueOf(circuitMap.get("customer_addr_a"))));
		    row.createCell(14).setCellValue("ToConfirm-A端业务设备");
		    row.createCell(15).setCellValue("ToConfirm-A端业务设备端口");
		    row.createCell(16).setCellValue("ToConfirm-A端移动电调联系人/电话");
		    
		    String aPersonAndPhoneInfo = Util.getValueOfStr(String.valueOf(circuitMap.get("customer_contact_person_a"))) + "/" + Util.getValueOfStr(String.valueOf(circuitMap.get("customer_contact_phone_a")));
		    row.createCell(17).setCellValue(aPersonAndPhoneInfo);

		    row.createCell(18).setCellValue(Util.getValueOfStr(String.valueOf(circuitMap.get("connection_site"))));
		    row.createCell(19).setCellValue(Util.getValueOfStr(String.valueOf(circuitMap.get("z_trans_site_name"))));
		    row.createCell(20).setCellValue(Util.getValueOfStr(String.valueOf(circuitMap.get("z_trans_room_name"))));
		    row.createCell(21).setCellValue(Util.getValueOfStr(String.valueOf(circuitMap.get("end_device_name"))));
		    row.createCell(22).setCellValue(Util.getValueOfStr(String.valueOf(circuitMap.get("end_equipport_name"))));
		    row.createCell(23).setCellValue(Util.getValueOfStr(String.valueOf(circuitMap.get("z_trans_cpt_name"))));
		   
		    
		    // TODO 25 26 
		    row.createCell(24).setCellValue(Util.getValueOfStr(String.valueOf(circuitMap.get("customer_addr_z"))));
		    row.createCell(25).setCellValue("ToConfirm-Z端业务设备");
		    row.createCell(26).setCellValue("ToConfirm-Z端业务设备端口");
		    String zPersonAndPhoneInfo = Util.getValueOfStr(String.valueOf(circuitMap.get("customer_contact_person_z"))) + "/" + Util.getValueOfStr(String.valueOf(circuitMap.get("customer_contact_phone_z")));
		    row.createCell(27).setCellValue(zPersonAndPhoneInfo);
		    
		    
		    // TODO 28
		    row.createCell(28).setCellValue("ToConfirm-电路路由");
		    
		    row.createCell(29).setCellValue(Util.getValueOfStr(String.valueOf(circuitMap.get("route_remark"))));
		    // TODO 需要映射
		    row.createCell(30).setCellValue(Util.getValueOfStr(String.valueOf(circuitMap.get("cic_level"))));
		    row.createCell(31).setCellValue(Util.getValueOfStr(String.valueOf(circuitMap.get("ext_ids"))));
		    
		    // 客户服务等级 和 业务保障等级 取同一字段 ,去掉客户服务等级,仅展示业务保障等级
		    // TODO 需要映射
		    row.createCell(32).setCellValue(Util.getValueOfStr(String.valueOf(circuitMap.get("service_level"))));
		    
		    row.createCell(33).setCellValue(Util.getValueOfStr(String.valueOf(circuitMap.get("handler_dept"))));
		    // TODO 需要映射
		    row.createCell(34).setCellValue(Util.getValueOfStr(String.valueOf(circuitMap.get("use_type"))));
		    row.createCell(35).setCellValue(Util.getValueOfStr(String.valueOf(circuitMap.get("is_priority_monitors_traph"))));
		   
		    row.createCell(36).setCellValue(Util.getValueOfStr(String.valueOf(circuitMap.get("rent_usdollars"))));
		    row.createCell(37).setCellValue(Util.getValueOfStr(String.valueOf(circuitMap.get("rent_yuan"))));
		    row.createCell(38).setCellValue(Util.getValueOfStr(String.valueOf(circuitMap.get("term"))));
		    row.createCell(39).setCellValue(Util.getValueOfStr(String.valueOf(circuitMap.get("circuit_alias"))));
		    row.createCell(40).setCellValue(Util.getValueOfStr(String.valueOf(circuitMap.get("comments"))));
         
		    // 自适应
            ExcelUtil.autoSizeColumn(sheet, row.getLastCellNum());
		}
	}
	
	/**
	 * 
	 * @param sheet
	 * @param style  首行样式
	 * @param height 首行高度
	 * @desc 设置首行的样式
	 */
	private void setCircuitExcelHeadRowStyle(XSSFSheet sheet, CellStyle style,Short height) {
		
		XSSFRow headerRow = sheet.createRow(0);
		headerRow.setHeight((short) height);
		
		Cell cell = headerRow.createCell(0);
		cell.setCellValue("序号");
		cell.setCellStyle(style);

		cell = headerRow.createCell(1);
		cell.setCellValue("ESOP订单号");
		cell.setCellStyle(style);
		
		cell = headerRow.createCell(2);
		cell.setCellValue("要求完成时间");
		cell.setCellStyle(style);
		
		cell = headerRow.createCell(3);
		cell.setCellValue("调度方式");
		cell.setCellStyle(style);
		
		cell = headerRow.createCell(4);
		cell.setCellValue("电路名称");
		cell.setCellStyle(style);
		
		cell = headerRow.createCell(5);
		cell.setCellValue("带宽");
		cell.setCellStyle(style);
		
		cell = headerRow.createCell(6);
		cell.setCellValue("速率");
		cell.setCellStyle(style);
		
		cell = headerRow.createCell(7);
		cell.setCellValue("传输QOS");
		cell.setCellStyle(style);
		
		cell = headerRow.createCell(8);
		cell.setCellValue("A端站点");
		cell.setCellStyle(style);
		
		cell = headerRow.createCell(9);
		cell.setCellValue("A端机房");
		cell.setCellStyle(style);
		
		cell = headerRow.createCell(10);
		cell.setCellValue("A端传输网元");
		cell.setCellStyle(style);
		
		cell = headerRow.createCell(11);
		cell.setCellValue("A端传输设备端口");
		cell.setCellStyle(style);
		
		cell = headerRow.createCell(12);
		cell.setCellValue("A端传输设备时隙");
		cell.setCellStyle(style);
		
		
		// 待确认14 15 TODO
		cell = headerRow.createCell(13);
		cell.setCellValue("A端客户地址");
		cell.setCellStyle(style);
		
		cell = headerRow.createCell(14);
		cell.setCellValue("A端业务设备");
		cell.setCellStyle(style);
		
		cell = headerRow.createCell(15);
		cell.setCellValue("A端业务设备端口");
		cell.setCellStyle(style);
		
		cell = headerRow.createCell(16);
		cell.setCellValue("A端移动电调联系人/电话");
		cell.setCellStyle(style);
		
		cell = headerRow.createCell(17);
		cell.setCellValue("A端客户联系人/电话");
		cell.setCellStyle(style);
		
		cell = headerRow.createCell(18);
		cell.setCellValue("转接站");
		cell.setCellStyle(style);
		
		cell = headerRow.createCell(19);
		cell.setCellValue("Z端站点");
		cell.setCellStyle(style);
		
		cell = headerRow.createCell(20);
		cell.setCellValue("Z端机房");
		cell.setCellStyle(style);
		
		cell = headerRow.createCell(21);
		cell.setCellValue("Z端传输网元");
		cell.setCellStyle(style);
		
		cell = headerRow.createCell(22);
		cell.setCellValue("Z端传输设备端口");
		cell.setCellStyle(style);
		
		cell = headerRow.createCell(23);
		cell.setCellValue("Z端传输设备时隙");
		cell.setCellStyle(style);
		
		// 待确认25 26
		cell = headerRow.createCell(24);
		cell.setCellValue("Z端客户地址");
		cell.setCellStyle(style);
		
		cell = headerRow.createCell(25);
		cell.setCellValue("Z端业务设备");
		cell.setCellStyle(style);
		
		cell = headerRow.createCell(26);
		cell.setCellValue("Z端业务设备端口");
		cell.setCellStyle(style);
		
		cell = headerRow.createCell(27);
		cell.setCellValue("Z端客户联系人/电话");
		cell.setCellStyle(style);
		
		// TODO 28
		cell = headerRow.createCell(28);
		cell.setCellValue("电路路由");
		cell.setCellStyle(style);
		
		
		cell = headerRow.createCell(29);
		cell.setCellValue("路由备注");
		cell.setCellStyle(style);
		
		cell = headerRow.createCell(30);
		cell.setCellValue("电路级别");
		cell.setCellStyle(style);
		
		cell = headerRow.createCell(31);
		cell.setCellValue("业务类型");
		cell.setCellStyle(style);
		
		
		// 客户服务等级 和 业务保障等级 取同一字段
		// headerRow.createCell(32).setCellValue("客户服务等级");
		cell = headerRow.createCell(32);
		cell.setCellValue("业务保障等级");
		cell.setCellStyle(style);
		
		cell = headerRow.createCell(33);
		cell.setCellValue("处理部门");
		cell.setCellStyle(style);
		
		cell = headerRow.createCell(34);
		cell.setCellValue("用途");
		cell.setCellStyle(style);
		
		cell = headerRow.createCell(35);
		cell.setCellValue("是否为重点监控电路");
		cell.setCellStyle(style);
		
		cell = headerRow.createCell(36);
		cell.setCellValue("租金(美元)");
		cell.setCellStyle(style);
		
		cell = headerRow.createCell(37);
		cell.setCellValue("租金(人民币)");
		cell.setCellStyle(style);
		
		cell = headerRow.createCell(38);
		cell.setCellValue("租期");
		cell.setCellStyle(style);
		
		cell = headerRow.createCell(39);
		cell.setCellValue("电路别名");
		cell.setCellStyle(style);
		
		cell = headerRow.createCell(40);
		cell.setCellValue("备注");
		cell.setCellStyle(style);
		
		
	}
	
	/**
	 * 
	 * @param sheet
	 * @param circuitInfoMap
	 * @Desc 生成工单信息sheet列
	 */
	private void generateWorkOrderInfo(XSSFWorkbook workbook, Map<String, Object> circuitInfoMap) {
		// 工单信息sheet
		XSSFSheet sheet = workbook.createSheet("工单信息");
		
		// 设置cell的宽度  Set the width (in units of 1/256th of a character width)
		sheet.setColumnWidth(0, 10 * 256);
		sheet.setColumnWidth(1, 30 * 256);
		sheet.setColumnWidth(2, 10 * 256);
		sheet.setColumnWidth(3, 30 * 256);
		sheet.setColumnWidth(4, 15 * 256);
		sheet.setColumnWidth(5, 30 * 256);
		
		// 第1行 工单主题  
		XSSFRow row = sheet.createRow(0);
		row.setHeight(new Short("600"));
		row.createCell(0).setCellValue("工单主题");
		row.createCell(1).setCellValue(Util.getValueOfStr(String.valueOf(circuitInfoMap.get("flow_title"))));

		// 合并Cell  起止行号 终止行号 起止列号  终止列号   0-based
		CellRangeAddress cellRangeAddress=new CellRangeAddress(0,0,1,5);
		sheet.addMergedRegion(cellRangeAddress);
		
		// 第2行 详细描述
		row = sheet.createRow(1);
		row.setHeight(new Short("600"));
		row.createCell(0).setCellValue("详细描述");
		row.createCell(1).setCellValue(Util.getValueOfStr(String.valueOf(circuitInfoMap.get("flow_desc"))));
		cellRangeAddress=new CellRangeAddress(1,1,1,5);
		sheet.addMergedRegion(cellRangeAddress);
		
		// 第3行 主送
		row = sheet.createRow(2);
		row.setHeight(new Short("600"));
		row.createCell(0).setCellValue("主送");
		row.createCell(1).setCellValue("调用的外部接口，省略了");
		cellRangeAddress=new CellRangeAddress(2,2,1,5);
		sheet.addMergedRegion(cellRangeAddress);
		
		// 第4行的数据
		row = sheet.createRow(3);
		row.setHeight(new Short("600"));
		row.createCell(0).setCellValue("调单文号");
		row.createCell(1).setCellValue(Util.getValueOfStr(String.valueOf(circuitInfoMap.get("tune_number"))));
		
		row.createCell(2).setCellValue("创建时间");
		row.createCell(3).setCellValue(Util.getValueOfDate((Date)circuitInfoMap.get("send_time"),null));
		
		row.createCell(4).setCellValue("要求完成时间");
		row.createCell(5).setCellValue(Util.getValueOfDate((Date)circuitInfoMap.get("limit_time"),null));
		
		// 第5行的数据
		row = sheet.createRow(4);
		row.setHeight(new Short("600"));
		row.createCell(0).setCellValue("工单编号");
		row.createCell(1).setCellValue(Util.getValueOfStr(String.valueOf(circuitInfoMap.get("flow_no"))));
		
		row.createCell(2).setCellValue("ESOP订单号");
		row.createCell(3).setCellValue(Util.getValueOfStr(String.valueOf(circuitInfoMap.get("esop_order_no"))));
		
		row.createCell(4).setCellValue("客户编号");
		row.createCell(5).setCellValue(Util.getValueOfStr(String.valueOf(circuitInfoMap.get("esop_customer_no"))));
		
		// 第6行的数据
		row = sheet.createRow(5);
		row.setHeight(new Short("600"));
		row.createCell(0).setCellValue("客户名称");
		row.createCell(1).setCellValue(Util.getValueOfStr(String.valueOf(circuitInfoMap.get("esop_customer_name"))));
		
		row.createCell(2).setCellValue("所属省份");
		row.createCell(3).setCellValue(Util.getValueOfStr(String.valueOf(circuitInfoMap.get("esop_province_name"))));
		
		row.createCell(4).setCellValue("客户服务等级");
		row.createCell(5).setCellValue(Util.getValueOfStr(String.valueOf(circuitInfoMap.get("esop_service_level"))));
		
		// 第7行的数据
		row = sheet.createRow(6);
		row.setHeight(new Short("600"));
		row.createCell(0).setCellValue("客户经理");
		row.createCell(1).setCellValue(Util.getValueOfStr(String.valueOf(circuitInfoMap.get("esop_manager"))));
		
		row.createCell(2).setCellValue("联系电话");
		row.createCell(3).setCellValue(Util.getValueOfStr(String.valueOf(circuitInfoMap.get("esop_phone_no"))));
		
		row.createCell(4).setCellValue("业务跨域类别");
		row.createCell(5).setCellValue(Util.getValueOfStr(String.valueOf(circuitInfoMap.get("esop_domain_level"))));
		
		// 第8行的数据
		row = sheet.createRow(7);
		row.setHeight(new Short("600"));
		row.createCell(0).setCellValue("设计人");
		row.createCell(1).setCellValue(Util.getValueOfStr(String.valueOf(circuitInfoMap.get("send_man"))));
		
		row.createCell(2).setCellValue("联系电话");
		row.createCell(3).setCellValue(Util.getValueOfStr(String.valueOf(circuitInfoMap.get("send_no"))));
		
		row.createCell(4).setCellValue("调单类型");// 1 普通  2 集客
		row.createCell(5).setCellValue(Util.getValueOfStr(String.valueOf(circuitInfoMap.get("scheduling_type"))));
	}
	
	
	
}
