package com.artisan.util;

import java.text.SimpleDateFormat;
import java.util.Date;


public class Util {


	public static String getValueOfStr(String str){
		return (str==""||str==null || "null".equals(str))?"":str;
	}
	public static String getValueOfDate(Date date,SimpleDateFormat simpleDateFormat){
		
		if(date == null){
			return "";
		}
		if(simpleDateFormat!=null){
			return simpleDateFormat.format(date);
		}else{
			SimpleDateFormat sf = new SimpleDateFormat("yyyy-MM-dd hh:mm:ss");
			return sf.format(date);
		}
	}
	
	public static String toDownloadString(String s) {
		StringBuffer sb = new StringBuffer();
		for (int i = 0; i < s.length(); i++) {
			char c = s.charAt(i);
			if (c >= 0 && c <= 255) {
				sb.append(c);
			} else {
				byte[] b;
				try {
					b = Character.toString(c).getBytes("utf-8");
				} catch (Exception ex) {
					System.out.println(ex);
					b = new byte[0];
				}
				for (int j = 0; j < b.length; j++) {
					int k = b[j];
					if (k < 0)
						k += 256;
					sb.append("%" + Integer.toHexString(k).toUpperCase());
				}
			}
		}
		return sb.toString();
	}

}
