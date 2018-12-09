package com.artisan.dao;

import java.util.List;
import java.util.Map;

import org.apache.ibatis.annotations.Mapper;

@Mapper
public interface CircuitBusinessinfoModelMapper {
    // 查询工单信息 - 导出用
    List<Map<String, Object>> selectFormInfoByFlowId(String flowId);
}