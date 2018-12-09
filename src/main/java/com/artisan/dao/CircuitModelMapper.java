package com.artisan.dao;

import java.util.List;
import java.util.Map;

import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;


@Mapper
public interface CircuitModelMapper {
   
    // 查询电路信息 -导出用
    List<Map<String, Object>> selectCircuitInfoByFlowIdAndAttempType(@Param("flowId") String flowId,@Param("attempType") String attempType);

}