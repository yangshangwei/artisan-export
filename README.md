# artisan-export

Spring Boot + Mybatis + poi3.10-FINAL + swagger + druid + postgresql 搭建的导出excel实例

其中Mybatis使用resultType="map"   来接收数据，service层使用 List<Map<String, Object>>来接口，省略了字段和属性值的映射，没有domain类。 

<select id="selectCircuitInfoByFlowIdAndAttempType"  resultType="map"  parameterType="java.lang.String">
			SELECT
				pcb.esop_order_no,
				pcb.limit_time,
				ptc.attemp_type,
				ptc.name,
				ptc.bandwidth,
				ptc.rate,
				ptc.qos,
				ptc.a_trans_site_name,
				ptc.a_trans_room_name,
				ptc.start_device_name,
				ptc.start_equipport_name,
				ptc.a_trans_cpt_name,
				ptc.customer_addr_a,
				ptc.customer_contact_person_a,
				ptc.customer_contact_phone_a,
				ptc.customer_addr_z,
				ptc.customer_contact_person_z,
				ptc.customer_contact_phone_z,
				ptc.connection_site,
				ptc.z_trans_site_name,
				ptc.z_trans_room_name,
				ptc.end_device_name,
				ptc.end_equipport_name,
				ptc.z_trans_cpt_name,
				ptc.route_remark,
				ptc.cic_level,
				ptc.ext_ids,
				ptc.service_level,
				ptc.handler_dept,
				ptc.use_type,
				ptc.is_priority_monitors_traph,
				ptc.rent_usdollars,
				ptc.rent_yuan,
				ptc.term,
				ptc.circuit_alias,
				ptc.comments 
			FROM
				proc_temp_circuit ptc
				LEFT JOIN proc_circuit_businessinfo pcb ON ptc.flow_id = pcb.flow_id 
			WHERE 
				ptc.flow_id = #{flowId}
			And 
				ptc.attemp_type = #{attempType}
			ORDER BY
				ptc.id desc
	</select>
  
  
  
  
  Excel的导出: 多sheet列， 多cell合并， 设置背景色等功能
